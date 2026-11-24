import { child, children, type XmlElement } from '../util/xml.js';
import type { GradientFill } from './model.js';

const isColor = (c: XmlElement) => /^a:(srgbClr|schemeClr|sysClr|prstClr|scrgbClr|hslClr)$/.test(c.name);

type ColorOf = (el: XmlElement | undefined) => { color: string; alpha: number } | undefined;

/**
 * A DrawingML a:gradFill with its stops' colours resolved by `color`, and the middle stop's colour
 * to stand in for it where one colour must do. Undefined when no stop has a colour.
 */
export function parseGradient(grad: XmlElement, color: ColorOf): { gradient: GradientFill; mid: { color: string; alpha: number } } | undefined {
  const stops = children(child(grad, 'a:gsLst'), 'a:gs')
    .map((gs) => ({ pos: (Number(gs.attrs.pos?.replace(/%$/, '')) || 0) / (gs.attrs.pos?.endsWith('%') ? 100 : 100000), c: color(children(gs).find(isColor)) }))
    .filter((g): g is { pos: number; c: { color: string; alpha: number } } => !!g.c)
    .sort((a, b) => a.pos - b.pos);
  if (!stops.length) return undefined;
  const lin = child(grad, 'a:lin');
  const path = child(grad, 'a:path');
  const gradient: GradientFill = { stops: stops.map((g) => ({ pos: g.pos, color: g.c.color, ...(g.c.alpha < 1 && { alpha: g.c.alpha }) })), angle: (Number(lin?.attrs.ang) || 0) / 60000 };
  if (lin?.attrs.scaled === '1' || lin?.attrs.scaled === 'true') gradient.scaled = true;
  if (path && !lin) {
    const kind = path.attrs.path;
    gradient.path = kind === 'circle' || kind === 'rect' ? kind : 'shape';
    const rect = child(path, 'a:fillToRect');
    const f = (n: string) => (Number(rect?.attrs[n]) || 0) / 100000;
    gradient.focus = { l: f('l'), t: f('t'), r: f('r'), b: f('b') };
  }
  return { gradient, mid: stops[Math.floor(stops.length / 2)].c };
}
