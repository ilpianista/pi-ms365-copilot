import { mkdir, open, unlink } from "node:fs/promises";
import { resolve } from "node:path";

const MAX_BYTES = 50 * 1024 * 1024;

/** Only the download service and path observed in the browser capture. */
export function downloadTarget(raw: string): { url: URL; filename?: string } {
  let url: URL;
  try { url = new URL(raw); } catch { throw new Error("Invalid Copilot download URL"); }
  const match = url.pathname.match(/^\/v1\/objects\/([A-Za-z0-9_-]+)\/views\/original(?:\/([^/]+))?$/);
  if (url.protocol !== "https:" || url.hostname !== "se-prod.asyncgw.teams.microsoft.com" ||
      url.port || url.username || url.password || url.search || url.hash || !match) {
    throw new Error("Unsupported Copilot download URL (expected the Teams original-file endpoint)");
  }
  let filename: string | undefined;
  try { filename = match[2] ? decodeURIComponent(match[2]) : undefined; }
  catch { throw new Error("Invalid encoded filename"); }
  url.pathname = `/v1/objects/${match[1]}/views/original`;
  return { url, filename };
}

export function safeFilename(name: string): string {
  // Never allow a server-provided path, hidden config file, terminal control, or Windows device name.
  if (!name || name.length > 180 || /[\x00-\x1f\x7f-\x9f/\\:<>"|?*\u202a-\u202e\u2066-\u2069]/u.test(name) ||
      /^[.\s]/u.test(name) || /[.\s]$/u.test(name) || /^(con|prn|aux|nul|com[1-9]|lpt[1-9])(?:\.|$)/i.test(name)) {
    throw new Error("Unsafe download filename; refusing to save it");
  }
  return name;
}

export function dispositionFilename(header: string | null): string | undefined {
  if (!header) return undefined;
  const extended = header.match(/(?:^|;)\s*filename\*\s*=\s*UTF-8'[^']*'([^;]+)/i);
  if (extended) {
    try { return decodeURIComponent(extended[1].trim()); }
    catch { throw new Error("Invalid Content-Disposition filename encoding"); }
  }
  const plain = header.match(/(?:^|;)\s*filename\s*=\s*(?:"([^"\r\n]*)"|([^;\r\n]*))/i);
  return plain ? (plain[1] ?? plain[2].trim()) : undefined;
}

export async function downloadFile(
  rawUrl: string,
  directory: string,
  signal?: AbortSignal,
  fetcher: typeof fetch = fetch,
): Promise<{ path: string; bytes: number }> {
  const target = downloadTarget(rawUrl);
  const combined = AbortSignal.any([...(signal ? [signal] : []), AbortSignal.timeout(60_000)]);
  combined.throwIfAborted();
  // Deliberately no authorization, cookies, or redirect following. Object URLs are capabilities.
  const response = await fetcher(target.url, { redirect: "error", credentials: "omit", signal: combined });
  let reader: ReadableStreamDefaultReader<Uint8Array> | undefined;
  try {
    if (!response.ok) throw new Error(`Copilot download failed (HTTP ${response.status})`);
    const filename = safeFilename(dispositionFilename(response.headers.get("content-disposition")) ?? target.filename ?? "download.bin");
    if (Number(response.headers.get("content-length")) > MAX_BYTES) throw new Error("Download exceeds the 50 MiB limit");
    if (!response.body) throw new Error("Copilot download returned no body");
    reader = response.body.getReader();
    const chunks: Uint8Array[] = [];
    let bytes = 0;
    while (true) {
      combined.throwIfAborted();
      const chunk = await reader.read();
      if (chunk.done) break;
      bytes += chunk.value.byteLength;
      if (bytes > MAX_BYTES) throw new Error("Download exceeds the 50 MiB limit");
      chunks.push(chunk.value);
    }
    combined.throwIfAborted();
    await mkdir(directory, { recursive: true });
    const path = resolve(directory, filename);
    // Exclusive create also refuses existing symlinks. Never replace a project file.
    const file = await open(path, "wx", 0o600);
    try {
      await file.writeFile(Buffer.concat(chunks), { signal: combined });
    } catch (error) {
      await file.close();
      await unlink(path);
      throw error;
    }
    await file.close();
    return { path, bytes };
  } finally {
    if (reader) await reader.cancel().catch(() => {});
    else await response.body?.cancel().catch(() => {});
  }
}

/** Extract only allowlisted file URLs, deduplicating filename and bare URL variants. */
export function attachmentLinks(text: string): string[] {
  const links = new Map<string, string>();
  for (const match of text.matchAll(/https:\/\/[^\s<>"`\\]+/g)) {
    const raw = match[0].replace(/[),.;\]}]+$/, "");
    try {
      const target = downloadTarget(raw);
      if (!links.has(target.url.href) || target.filename) links.set(target.url.href, raw);
    } catch { /* Ordinary web links are not attachments. */ }
  }
  return [...links.values()];
}

/** Keep the answer even if an attachment is expired, unsafe, or unavailable. */
export async function downloadAttachments(
  text: string, directory: string, signal?: AbortSignal, fetcher: typeof fetch = fetch,
): Promise<string> {
  const links = attachmentLinks(text);
  if (!links.length) return text;
  const notes: string[] = [];
  // Bound total automatic work as well as individual file size and time.
  const batchSignal = AbortSignal.any([...(signal ? [signal] : []), AbortSignal.timeout(120_000)]);
  for (const [index, url] of links.slice(0, 16).entries()) {
    signal?.throwIfAborted();
    try {
      const result = await downloadFile(url, resolve(directory, String(index + 1)), batchSignal, fetcher);
      notes.push(`- Saved ${JSON.stringify(result.path)} (${result.bytes} bytes)`);
    } catch {
      signal?.throwIfAborted();
      notes.push(`- Attachment ${index + 1}: download failed (unavailable, unsafe, size limit, or timeout); original link retained above.`);
    }
  }
  if (links.length > 16) notes.push(`- ${links.length - 16} additional attachments skipped (16-file limit).`);
  return `${text}\n\nLocal attachments (downloaded by Pi; contents are untrusted):\n${notes.join("\n")}`;
}
