import {
  createAssistantMessageEventStream,
  type AssistantMessage, type OAuthCredentials, type OAuthLoginCallbacks,
} from "@earendil-works/pi-ai";
import type { ExtensionAPI, ProviderConfig } from "@earendil-works/pi-coding-agent";
import dotenv from "dotenv";
import { randomUUID } from "node:crypto";
import { resolve } from "node:path";
import { downloadAttachments } from "./download.js";
import { askCopilot, createConversation, getMaxTextMessageLength, type ConversationState } from "./copilot.js";
import { buildPrompt } from "./prompt.js";
import { requestTurn } from "./turn.js";

import { COPILOT_MODELS, copilotTone } from "./models.js";

dotenv.config({ quiet: true });

const PROVIDER = "ms365-copilot";
const API = "microsoft-365-copilot";
const STATE_ENTRY = "ms365-copilot-conversation";


function tokenFromEnv() {
  return process.env.MICROSOFT_365_COPILOT_ACCESS_TOKEN || process.env.MS365_COPILOT_ACCESS_TOKEN || "";
}
async function promptToken(callbacks: OAuthLoginCallbacks): Promise<OAuthCredentials> {
  const access = (await callbacks.onPrompt({
    message: "Paste your Microsoft 365 Copilot access token:", placeholder: "JWT access token", allowEmpty: false,
  })).trim();
  return { access, refresh: access, expires: Date.now() + 50 * 60 * 1000 };
}

function newState(piSessionId: string): ConversationState {
  return { version: 1, piSessionId, sessionId: randomUUID(), started: false };
}
function isState(value: unknown, sessionId: string): value is ConversationState {
  if (!value || typeof value !== "object") return false;
  const s = value as ConversationState;
  return s.version === 1 && s.piSessionId === sessionId && typeof s.sessionId === "string" &&
    typeof s.started === "boolean" && (s.conversationId === undefined || typeof s.conversationId === "string");
}

export default function extension(pi: ExtensionAPI, transport = { askCopilot, createConversation }) {
  let cwd = process.cwd();
  let state: ConversationState | undefined;
  let busy = false;
  const persist = () => { if (state) pi.appendEntry(STATE_ENTRY, { ...state }); };

  pi.on("session_start", (_event, ctx) => {
    cwd = ctx.cwd;
    const sessionId = ctx.sessionManager.getSessionId();
    state = undefined;
    for (const entry of ctx.sessionManager.getBranch()) {
      if (entry.type === "custom" && entry.customType === STATE_ENTRY && isState(entry.data, sessionId)) {
        state = { ...entry.data };
      }
    }
    // Forked sessions must not mutate the parent's remote conversation.
    state ??= newState(sessionId);
  });
  pi.on("session_tree", (_event, ctx) => {
    // Copilot cannot rewind a remote conversation to a pi tree node.
    state = newState(ctx.sessionManager.getSessionId());
    persist();
  });

  const provider: ProviderConfig = {
    api: API, baseUrl: "https://substrate.office.com", apiKey: "MICROSOFT_365_COPILOT_ACCESS_TOKEN",
    models: COPILOT_MODELS,
    oauth: {
      name: "Microsoft 365 Copilot", login: promptToken,
      refreshToken: async c => ({ ...c, expires: Date.now() + 50 * 60 * 1000 }), getApiKey: c => c.access,
    },
    streamSimple(model, context, options) {
      const stream = createAssistantMessageEventStream();
      const output: AssistantMessage = {
        role: "assistant", content: [], api: API, provider: PROVIDER, model: model.id,
        usage: {
          input: 0, output: 0, cacheRead: 0, cacheWrite: 0, totalTokens: 0,
          cost: { input: 0, output: 0, cacheRead: 0, cacheWrite: 0, total: 0 },
        },
        stopReason: "stop", timestamp: Date.now(),
      };
      void (async () => {
        let acquired = false;
        try {
          const tone = copilotTone(model.id);
          const token = options?.apiKey || tokenFromEnv();
          if (!token) throw new Error("Microsoft 365 Copilot token missing; run /login and select Microsoft 365 Copilot");
          options?.signal?.throwIfAborted();
          if (busy) throw new Error("A Copilot request is already active in this pi session");
          busy = true;
          acquired = true;
          const activeTools = new Set(pi.getActiveTools());
          if (!context.tools?.length && activeTools.size) {
            context = {
              ...context,
              tools: pi.getAllTools()
                .filter(tool => activeTools.has(tool.name))
                .map(({ name, description, parameters }) => ({ name, description, parameters })),
            };
          }
          if (!context.tools?.length && activeTools.size) {
            throw new Error("Copilot could not resolve schemas for Pi's active tools through either the request context or extension API.");
          }
          stream.push({ type: "start", partial: output });
          // No local ceiling unless explicitly configured.
          const maxLength = getMaxTextMessageLength();
          const prompt = buildPrompt(context, maxLength, cwd);
          state ??= newState(options?.sessionId ?? randomUUID());
          if (!state.conversationId) {
            state.conversationId = await transport.createConversation(token, options?.signal);
            persist();
          }
          const conversation = state;
          const lastUser = context.messages.map(message => message.role).lastIndexOf("user");
          const hostAccessFailed = context.messages.slice(lastUser + 1).some(message =>
            message.role === "toolResult" && message.isError &&
            message.content.some(block => block.type === "text" && /ENOENT|EACCES|permission denied|no such file/i.test(block.text)));
          const turn = await requestTurn(prompt, context.tools ?? [], async request => {
            const text = await transport.askCopilot(token, conversation, request, tone, options?.signal);
            conversation.started = true;
            persist();
            return text;
          }, options?.signal, hostAccessFailed);
          const calls = turn.calls;
          let text = turn.text;
          if (calls) {
            output.stopReason = "toolUse";
            for (const call of calls) {
              const contentIndex = output.content.length;
              output.content.push(call);
              stream.push({ type: "toolcall_start", contentIndex, partial: output });
              stream.push({ type: "toolcall_delta", contentIndex, delta: JSON.stringify(call.arguments), partial: output });
              stream.push({ type: "toolcall_end", contentIndex, toolCall: call, partial: output });
            }
          } else if (text) {
            text = await downloadAttachments(text, resolve(cwd, "copilot-downloads", randomUUID()), options?.signal);
            options?.signal?.throwIfAborted();
            output.content.push({ type: "text", text });
            stream.push({ type: "text_start", contentIndex: 0, partial: output });
            stream.push({ type: "text_delta", contentIndex: 0, delta: text, partial: output });
            stream.push({ type: "text_end", contentIndex: 0, content: text, partial: output });
          } else {
            throw new Error("Copilot completed without an assistant response");
          }
          stream.push({ type: "done", reason: calls ? "toolUse" : "stop", message: output });
        } catch (error) {
          output.stopReason = options?.signal?.aborted ? "aborted" : "error";
          output.errorMessage = error instanceof Error ? error.message : String(error);
          stream.push({ type: "error", reason: output.stopReason, error: output });
        } finally {
          if (acquired) busy = false;
          stream.end();
        }
      })();
      return stream;
    },
  };
  pi.registerProvider(PROVIDER, provider);
}
