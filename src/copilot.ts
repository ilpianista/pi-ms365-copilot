import WebSocket from "ws";
import { randomUUID } from "node:crypto";
import type { CopilotTone } from "./models.js";

const BASE = "https://substrate.office.com";
export interface ConversationState {
  version: 1;
  piSessionId: string;
  conversationId?: string;
  sessionId: string;
  started: boolean;
}

export function tokenClaims(token: string): { tid?: string; oid?: string } {
  try { return JSON.parse(Buffer.from(token.split(".")[1], "base64url").toString()); }
  catch { return {}; }
}

function identity(token: string) {
  const { tid, oid } = tokenClaims(token);
  if (!tid || !oid) throw new Error("The token must contain tid and oid claims");
  return { tid, oid };
}

class CopilotHttpError extends Error {
  constructor(path: string, readonly status: number) {
    super(`Copilot ${path} failed (HTTP ${status})`);
  }
}

async function request(token: string, path: string, method: string, signal?: AbortSignal) {
  signal?.throwIfAborted();
  const { oid } = identity(token);
  const response = await fetch(`${BASE}${path}`, {
    method,
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
      "X-RoutingParameter-SessionKey": oid,
      "client-request-id": randomUUID(),
    },
    ...(method === "POST" ? { body: JSON.stringify({ source: "officeweb" }) } : {}),
    signal: AbortSignal.any([...(signal ? [signal] : []), AbortSignal.timeout(30_000)]),
  });
  if (!response.ok) throw new CopilotHttpError(path, response.status);
  return response.json();
}

/** Local prompt ceiling, not a discovered or verified server limit. */
export function getMaxTextMessageLength(): number {
  const limit = Number(process.env.MS365_COPILOT_MAX_TEXT_MESSAGE_LENGTH ?? 6944);
  if (!Number.isSafeInteger(limit) || limit <= 0) {
    throw new Error("MS365_COPILOT_MAX_TEXT_MESSAGE_LENGTH must be a positive safe integer");
  }
  return limit;
}

export async function createConversation(token: string, signal?: AbortSignal): Promise<string> {
  let data;
  try {
    data = await request(token, "/m365Copilot/Conversation", "POST", signal);
  } catch (error) {
    signal?.throwIfAborted();
    if (!(error instanceof CopilotHttpError) || error.status !== 404) throw error;
    // Match the standalone client's WebSocket bootstrap with a fresh local ID.
    return randomUUID();
  }
  if (data.result?.value && data.result.value !== "Success") {
    throw new Error(`Copilot conversation creation failed: ${data.result.value}`);
  }
  const id = data.conversationId ?? data.conversation?.conversationId;
  if (typeof id !== "string" || !id.trim()) throw new Error("Copilot did not return a conversationId");
  return id;
}

export function chatPayload(text: string, tone: CopilotTone, state: ConversationState, requestId: string) {
  const clientInfo = {
    clientPlatform: "ccmcopilot-web", clientAppName: "Office", clientEntrypoint: "ccmcopilot-officeweb",
    clientSessionId: state.sessionId, ProductCategory: "Chat", clientAppType: "Web",
    productEntryPoint: "ChatPanel", deviceOS: "Linux", deviceType: "Desktop", clientPlatformVersion: "10",
  };
  return {
    arguments: [{
      source: "officeweb", clientCorrelationId: requestId, sessionId: state.sessionId,
      conversationId: state.conversationId,
      optionsSets: ["cwc_flux_v3", "rich_responses", "enable_search_result_progress_messages"],
      options: {}, extraExtensionParameters: {},
      allowedMessageTypes: ["Chat", "Progress", "EndOfRequest", "ReferencesListComplete"],
      requestId, traceId: requestId, isStartOfSession: !state.started, clientInfo,
      message: {
        author: "user", inputMethod: "Keyboard", text, requestId, messageType: "Chat", locale: "en-gb",
        locationInfo: { timeZone: "UTC", timeZoneOffset: 0 }, clientInfo,
      },
      plugins: [{ Id: "BingWebSearch", Source: "BuiltIn" }],
      tone, streamingMode: "ConciseWithPadding", disconnectBehavior: "continue",
    }],
    invocationId: "0", target: "chat", type: 4,
  };
}

/** Buffer snapshots until completion: never expose partial tool JSON as ordinary assistant text. */
export async function askCopilot(token: string, state: ConversationState, text: string, tone: CopilotTone, signal?: AbortSignal): Promise<string> {
  signal?.throwIfAborted();
  const { tid, oid } = identity(token);
  if (!state.conversationId) throw new Error("Copilot conversation has not been bootstrapped");
  const requestId = randomUUID();
  const url = new URL(`${BASE.replace("https:", "wss:")}/m365Copilot/Chathub/${oid}@${tid}`);
  for (const [key, value] of Object.entries({
    chatsessionid: requestId, XRoutingParameterSessionKey: oid, clientrequestid: requestId,
    "X-SessionId": state.sessionId, ConversationId: state.conversationId, access_token: token,
    source: '"officeweb"', product: "Office", agentHost: "Bizchat.FullScreen", licenseType: "Starter",
    isEdu: "false", agent: "web", scenario: "OfficeWebIncludedCopilot",
  })) url.searchParams.set(key, value);
  const ws = new WebSocket(url, {
    headers: { Origin: "https://copilot.cloud.microsoft", "User-Agent": process.env.MS365_COPILOT_USER_AGENT || "Mozilla/5.0" },
    handshakeTimeout: 30_000,
  });
  return new Promise<string>((resolve, reject) => {
    let done = false, buffer = "", invoked = false;
    const texts = new Map<string, string>();
    const timer = setTimeout(() => finish(new Error("Copilot request timed out")), 180_000);
    const abort = () => finish(new Error("Request aborted"));
    function finish(error?: Error) {
      if (done) return;
      done = true;
      clearTimeout(timer);
      signal?.removeEventListener("abort", abort);
      if (error) reject(error);
      else resolve([...texts.values()].join("\n\n"));
      if (ws.readyState === WebSocket.OPEN) ws.close();
      else if (ws.readyState === WebSocket.CONNECTING) ws.terminate();
    }
    const send = (value: unknown) => ws.send(JSON.stringify(value) + "\x1e");
    ws.on("open", () => { if (!done) send({ protocol: "json", version: 1 }); });
    ws.on("message", raw => {
      if (done) return;
      buffer += raw.toString();
      const frames = buffer.split("\x1e");
      buffer = frames.pop() ?? "";
      for (const frame of frames) {
        if (!frame || done) continue;
        try {
          const message = JSON.parse(frame);
          if (message.error) throw new Error(`Copilot: ${message.error}`);
          if (!invoked && message.type === undefined) {
            invoked = true;
            send(chatPayload(text, tone, state, requestId));
            continue;
          }
          const result = message.item?.result;
          if (result?.value && result.value !== "Success") throw new Error(`Copilot: ${result.value}`);
          for (const item of [...(message.arguments?.[0]?.messages ?? []), ...(message.item?.messages ?? [])]) {
            if (item.author === "bot" && (!item.messageType || item.messageType === "Chat") && typeof item.text === "string") {
              texts.set(item.messageId ?? item.id ?? "answer", item.text);
            }
          }
          if (message.type === 6) send({ type: 6 });
          if (message.type === 7) throw new Error("Copilot closed the connection before completion");
          if (message.type === 2 || message.type === 3 || message.item?.turnState === "Completed") finish();
        } catch (error) { finish(error instanceof Error ? error : new Error(String(error))); }
      }
    });
    ws.on("unexpected-response", (_req, response) => {
      response.resume();
      finish(new Error(`Copilot WebSocket upgrade failed (HTTP ${response.statusCode})`));
    });
    ws.on("error", error => finish(error));
    ws.on("close", code => { if (!done) finish(new Error(`Copilot socket closed before completion (${code})`)); });
    signal?.addEventListener("abort", abort, { once: true });
    if (signal?.aborted) abort();
  });
}
