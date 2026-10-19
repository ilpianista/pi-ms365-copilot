import type { Tool, ToolCall } from "@earendil-works/pi-ai";
import { parseToolCalls } from "./prompt.js";

export type CopilotTurn = { calls: ToolCall[]; text?: never } | { text: string; calls?: never };

export function parseTurn(text: string, tools: Tool[]): CopilotTurn {
  if (!tools.length) {
    if (!text.trim()) throw new Error("Copilot completed without an assistant response");
    return { text };
  }
  const calls = parseToolCalls(text, tools);
  if (calls) return { calls };
  const trimmed = text.trim();
  const body = trimmed.match(/^```(?:json)?\s*\n([\s\S]*?)\n```$/)?.[1] ?? trimmed;
  let value;
  try { value = JSON.parse(body); } catch {}
  if (value && Object.keys(value).length === 1 && typeof value.pi_final === "string" && value.pi_final.trim()) {
    return { text: value.pi_final };
  }
  throw new Error("Expected a pi_tool_calls or pi_final JSON envelope");
}

export async function requestTurn(
  prompt: string,
  tools: Tool[],
  ask: (prompt: string) => Promise<string>,
  signal?: AbortSignal,
  hostAccessFailed = false,
): Promise<CopilotTurn> {
  let nextPrompt = prompt;
  for (let attempt = 0; attempt < 3; attempt++) {
    signal?.throwIfAborted();
    const text = await ask(nextPrompt);
    signal?.throwIfAborted();
    try {
      const turn = parseTurn(text, tools);
      if (tools.length && turn.text && !hostAccessFailed &&
          /(?:can(?:not|['’]t)|unable to) access (?:the |your )?(?:repo(?:sitory)?|workspace)|(?:paths?|workspace|repository).{0,100}(?:not mounted|unavailable|inaccessible)|(?:attach|upload).{0,60}(?:patch\.patch|repository files)/is.test(turn.text)) {
        throw new Error("Copilot claimed the host workspace was unavailable without a failed Pi tool result");
      }
      return turn;
    } catch (error) {
      if (!tools.length || attempt === 2) {
        throw new Error(`Copilot tool protocol failed: ${error instanceof Error ? error.message : String(error)}. No host tools were executed from this response. Copilot must request Pi tools; its remote sandbox cannot access your repository.`);
      }
      nextPrompt = 'Your last response was not a valid Pi response. Return ONLY one JSON object: {"pi_tool_calls":[{"name":"enabled_tool_name","arguments":{}}]} to request host tools, or {"pi_final":"answer"} if finished. Use the enabled tools and schemas from the preceding request. To inspect or change repository files, request host tools now, not Copilot sandbox tools. A plan or promise does not execute tools. Do not claim host files are unavailable based on your remote sandbox. Do not repeat an action already confirmed by a toolResult.';
    }
  }
  throw new Error("Copilot tool protocol failed");
}
