export const ACCEPTED_TYPES = [
  "application/pdf",
  "image/png",
  "image/jpeg",
  "image/webp",
  "image/heic",
  "image/heif",
];

export const MAX_FILE_SIZE = 35 * 1024 * 1024;

export function isAcceptedFile(file: File) {
  const byMime = ACCEPTED_TYPES.includes(file.type);
  const ext = file.name.toLowerCase().split(".").pop();
  const byExt = ["pdf", "png", "jpg", "jpeg", "webp", "heic", "heif"].includes(ext ?? "");
  return byMime || byExt;
}

export function formatBytes(bytes: number) {
  if (!Number.isFinite(bytes) || bytes <= 0) return "0 B";
  const units = ["B", "KB", "MB", "GB"];
  const i = Math.min(Math.floor(Math.log(bytes) / Math.log(1024)), units.length - 1);
  return `${(bytes / 1024 ** i).toFixed(i === 0 ? 0 : 1)} ${units[i]}`;
}

export function outputName(input: string) {
  const base = input.replace(/\.[^/.]+$/, "").replace(/[^a-zA-Z0-9À-ÿ _-]/g, "").trim() || "document";
  return `${base}_PX_editable.docx`;
}

export function downloadBlob(blob: Blob, fileName: string) {
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = fileName;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  setTimeout(() => URL.revokeObjectURL(url), 2000);
}
