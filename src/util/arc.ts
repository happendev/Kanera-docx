/**
 * Cubic Béziers along an elliptical arc from DrawingML's start angle through its sweep (degrees,
 * clockwise). The angles are visual: a ray at that angle meets the ellipse where the arc starts.
 */
export function arc(cx: number, cy: number, rx: number, ry: number, start: number, sweep: number): { from: number[]; segs: Array<{ op: 'C'; pts: number[] }> } {
  const param = (deg: number) => Math.atan2(rx * Math.sin((deg * Math.PI) / 180), ry * Math.cos((deg * Math.PI) / 180));
  const t0 = param(start);
  let dt = param(start + sweep) - t0;
  if (sweep > 0) while (dt <= 1e-9) dt += 2 * Math.PI;
  if (sweep < 0) while (dt >= -1e-9) dt -= 2 * Math.PI;
  const n = Math.max(1, Math.ceil(Math.abs(dt) / (Math.PI / 2) - 1e-9));
  const step = dt / n;
  const k = (4 / 3) * Math.tan(step / 4);
  const pt = (t: number) => [cx + rx * Math.cos(t), cy + ry * Math.sin(t)];
  const segs: Array<{ op: 'C'; pts: number[] }> = [];
  for (let i = 0; i < n; i++) {
    const a = t0 + i * step;
    const b = a + step;
    segs.push({ op: 'C', pts: [cx + rx * (Math.cos(a) - k * Math.sin(a)), cy + ry * (Math.sin(a) + k * Math.cos(a)), cx + rx * (Math.cos(b) + k * Math.sin(b)), cy + ry * (Math.sin(b) - k * Math.cos(b)), ...pt(b)] });
  }
  return { from: pt(t0), segs };
}
