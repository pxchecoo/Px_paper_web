export function rgbToHex(r: number, g: number, b: number) {
  const normalized = [r, g, b].map((value) => {
    const v = value <= 1 ? Math.round(value * 255) : Math.round(value);
    return Math.max(0, Math.min(255, v));
  });
  return normalized.map((v) => v.toString(16).padStart(2, "0")).join("").toUpperCase();
}

export function cmykToRgbHex(c: number, m: number, y: number, k: number) {
  const normalize = (v: number) => (v > 1 ? v / 100 : v);
  const [cc, mm, yy, kk] = [c, m, y, k].map(normalize);
  return rgbToHex(255 * (1 - cc) * (1 - kk), 255 * (1 - mm) * (1 - kk), 255 * (1 - yy) * (1 - kk));
}

export function estimateTextColor(
  ctx: CanvasRenderingContext2D,
  bbox: { x0: number; y0: number; x1: number; y1: number },
) {
  const x = Math.max(0, Math.floor(bbox.x0));
  const y = Math.max(0, Math.floor(bbox.y0));
  const width = Math.max(1, Math.min(ctx.canvas.width - x, Math.ceil(bbox.x1 - bbox.x0)));
  const height = Math.max(1, Math.min(ctx.canvas.height - y, Math.ceil(bbox.y1 - bbox.y0)));
  if (width <= 0 || height <= 0) return "111827";

  const data = ctx.getImageData(x, y, width, height).data;
  const pixels: Array<[number, number, number]> = [];
  const stride = Math.max(1, Math.floor(Math.sqrt((width * height) / 1000)));

  for (let py = 0; py < height; py += stride) {
    for (let px = 0; px < width; px += stride) {
      const i = (py * width + px) * 4;
      const r = data[i];
      const g = data[i + 1];
      const b = data[i + 2];
      const a = data[i + 3];
      if (a < 80) continue;
      const distanceFromWhite = Math.sqrt((255 - r) ** 2 + (255 - g) ** 2 + (255 - b) ** 2);
      const max = Math.max(r, g, b);
      const min = Math.min(r, g, b);
      const saturation = max - min;
      const luminance = 0.2126 * r + 0.7152 * g + 0.0722 * b;
      if (distanceFromWhite > 70 && (luminance < 215 || saturation > 45)) pixels.push([r, g, b]);
    }
  }

  if (!pixels.length) return "111827";
  pixels.sort((a, b) => {
    const la = 0.2126 * a[0] + 0.7152 * a[1] + 0.0722 * a[2];
    const lb = 0.2126 * b[0] + 0.7152 * b[1] + 0.0722 * b[2];
    return la - lb;
  });
  const sample = pixels.slice(0, Math.max(1, Math.ceil(pixels.length * 0.6)));
  const avg = sample.reduce(
    (acc, p) => [acc[0] + p[0], acc[1] + p[1], acc[2] + p[2]],
    [0, 0, 0],
  );
  return rgbToHex(avg[0] / sample.length, avg[1] / sample.length, avg[2] / sample.length);
}
