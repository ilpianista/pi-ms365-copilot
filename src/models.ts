import type { ProviderConfig } from "@mariozechner/pi-coding-agent";

// Routing values observed in the browser, not verified underlying model identities.
const routes = [
  { id: "copilot", name: "Auto", tone: "Magic", reasoning: false },
  { id: "copilot-quick", name: "Quick", tone: "Chat", reasoning: false },
  { id: "copilot-deeper", name: "Think Deeper", tone: "Reasoning", reasoning: true },
  { id: "copilot-sol-quick", name: "SOL Quick", tone: "Gpt_5_6_Chat", reasoning: false },
  { id: "copilot-sol-think", name: "SOL Think", tone: "Gpt_5_6_Reasoning", reasoning: true },
] as const;

export type CopilotTone = typeof routes[number]["tone"];

export function copilotTone(modelId: string): CopilotTone {
  const route = routes.find(route => route.id === modelId);
  if (!route) throw new Error(`Unknown Microsoft 365 Copilot model: ${modelId}`);
  return route.tone;
}

export const COPILOT_MODELS: NonNullable<ProviderConfig["models"]> = routes.map(route => ({
  id: route.id, name: `Microsoft 365 Copilot — ${route.name}`, reasoning: route.reasoning,
  // Reasoning tones are fixed: no off switch or adjustable effort was observed.
  ...(route.reasoning ? { thinkingLevelMap: {
    off: null, minimal: null, low: null, medium: null, high: route.tone, xhigh: null,
  } } : {}),
  input: ["text"],
  cost: { input: 0, output: 0, cacheRead: 0, cacheWrite: 0 }, contextWindow: 128000, maxTokens: 8192,
}));
