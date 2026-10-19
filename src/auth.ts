import type { OAuthCredentials, OAuthLoginCallbacks } from "@earendil-works/pi-ai";

/** Read expiration metadata only; decoding does not verify a JWT's signature. */
export function tokenExpiration(token: string): number | undefined {
  const parts = token.split(".");
  if (parts.length !== 3) return undefined;
  try {
    const claims: unknown = JSON.parse(Buffer.from(parts[1], "base64url").toString("utf8"));
    if (!claims || typeof claims !== "object" || !("exp" in claims)) return undefined;
    const exp = claims.exp;
    if (typeof exp !== "number" || !Number.isFinite(exp)) return undefined;
    const milliseconds = exp * 1000;
    return Number.isFinite(new Date(milliseconds).getTime()) ? milliseconds : undefined;
  } catch {
    return undefined;
  }
}

export async function promptToken(callbacks: OAuthLoginCallbacks): Promise<OAuthCredentials> {
  const access = (await callbacks.onPrompt({
    message: "Paste your Microsoft 365 Copilot access token:", placeholder: "JWT access token", allowEmpty: false,
  })).trim();
  callbacks.signal?.throwIfAborted();
  if (!access) throw new Error("An access token is required.");
  const now = Date.now();
  const expiration = tokenExpiration(access);
  if (expiration === undefined) {
    throw new Error("Cannot read the token expiration. Paste a JWT access token with a valid exp claim.");
  }
  if (expiration <= now) {
    throw new Error(`Access token expired at ${new Date(expiration).toISOString()}. Extract a fresh token and run /login again.`);
  }
  callbacks.onProgress?.(`Access token expires at ${new Date(expiration).toISOString()} (UTC; according to its JWT exp claim). Pasting it does not extend its lifetime.`);
  return { access, refresh: access, expires: expiration };
}
