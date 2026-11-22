import { arc } from '../util/arc.js';
import { child, children, descendant, descendants, path, resolveAlternateContent, type XmlElement } from '../util/xml.js';
import { parseChart, type ChartPointBudget } from './chart.js';
import { parseGradient } from './gradient.js';
import type { AnchorProps, Block, Chart, Color, Drawing, Fill, ImageRef, LineStyle, PositionSpec, TextBox } from './model.js';
import type { OpcPackage } from './package.js';
import { themeColorValue, type Theme } from './theme.js';
import { applyColorMods, emu, hexColor, measure, num } from './values.js';

export interface DrawingContext {
  pkg: OpcPackage;
  part: string;
  theme: Theme;
  colorMapping?: Record<string, string>;
  parseBlocks(container: XmlElement): Block[];
  images: Map<string, ImageRef>;
  chartBudget: ChartPointBudget;
  charts: Map<string, Chart | undefined>;
  /** Record pictures by part name without reading their bytes (text extraction). */
  skipImageData?: boolean;
}

const NO_DATA = new Uint8Array(0);

const PRESET_COLORS: Record<string, string> = { black: '000000', white: 'FFFFFF', red: 'FF0000', green: '008000', blue: '0000FF', yellow: 'FFFF00', gray: '808080' };

export function loadImage(ctx: DrawingContext, rid: string | undefined): ImageRef | undefined {
  if (!rid) return undefined;
  const rel = ctx.pkg.rels(ctx.part).get(rid);
  if (!rel || rel.external) return undefined;
  const cached = ctx.images.get(rel.target);
  if (cached) return cached;
  const data = ctx.skipImageData ? (ctx.pkg.zip.has(rel.target) ? NO_DATA : null) : ctx.pkg.bytes(rel.target);
  if (!data) return undefined;
  const ref: ImageRef = { id: rel.target, data };
  ctx.images.set(rel.target, ref);
  return ref;
}

function loadBlip(ctx: DrawingContext, blip: XmlElement | undefined): ImageRef | undefined {
  const fallback = loadImage(ctx, blip?.attrs['r:embed']);
  const svgBlip = children(child(blip, 'a:extLst'), 'a:ext').map((ext) => child(ext, 'asvg:svgBlip')).find(Boolean);
  const svg = loadImage(ctx, svgBlip?.attrs['r:embed']);
  const ref = svg && fallback ? { ...fallback, svg } : svg ?? fallback;
  // Brightness and contrast in thousandths of a percent, or "70%" in strict documents. Washout is
  // bright 70% with contrast -70%.
  const lum = child(blip, 'a:lum');
  const amount = (v: string | undefined) => (v?.endsWith('%') ? (num(v.slice(0, -1)) ?? 0) / 100 : (num(v) ?? 0) / 100000);
  const contrast = Math.min(0.99, Math.max(-1, amount(lum?.attrs.contrast)));
  return adjusted(ref, contrast <= 0 ? 1 + contrast : 1 / (1 - contrast), amount(lum?.attrs.bright));
}

/**
 * A picture recoloured by its brightness and contrast settings. Measured against Word's PDFs, each
 * colour component c (0–1) is drawn as s·(c − ½) + ½ + b·(1 + s)/2, clipped to [0, 1], where s is
 * the contrast's slope and b the brightness.
 */
function adjusted(ref: ImageRef | undefined, slope: number, bright: number): ImageRef | undefined {
  if (!ref || (Math.abs(slope - 1) < 1e-4 && Math.abs(bright) < 1e-4)) return ref;
  const r4 = (v: number) => Math.round(v * 10000) / 10000;
  const adjust = { slope: r4(slope), offset: r4(0.5 - slope / 2 + (bright * (1 + slope)) / 2) };
  // Its own id: the PDF embeds each id once, and the same picture may also appear unadjusted.
  return { ...ref, id: `${ref.id}#${adjust.slope},${adjust.offset}`, adjust };
}

/** Parse a DrawingML colour element (srgbClr, schemeClr, ...) including its transforms. */
export function parseColor(el: XmlElement | undefined, ctx: DrawingContext, phClr?: string): { color: string; alpha: number } | undefined {
  if (!el) return undefined;
  let base: string | undefined;
  switch (el.name) {
    case 'a:srgbClr':
      base = el.attrs.val;
      break;
    case 'a:schemeClr':
      base = el.attrs.val === 'phClr' && phClr ? phClr : themeColorValue(ctx.theme, el.attrs.val ?? '', ctx.colorMapping);
      break;
    case 'a:sysClr':
      base = el.attrs.lastClr ?? (el.attrs.val === 'window' ? 'FFFFFF' : '000000');
      break;
    case 'a:prstClr':
      base = PRESET_COLORS[el.attrs.val ?? ''] ?? '000000';
      break;
    case 'a:scrgbClr': {
      const c = (v: string | undefined) => Math.round(((v?.endsWith('%') ? (num(v.slice(0, -1)) ?? 0) * 1000 : (num(v) ?? 0)) / 100000) * 255);
      base = [c(el.attrs.r), c(el.attrs.g), c(el.attrs.b)].map((v) => v.toString(16).padStart(2, '0')).join('');
      break;
    }
    default:
      return undefined;
  }
  if (!base) return undefined;
  // Strict documents write percentages as "15%" rather than thousandths of a percent.
  const amount = (v: string | undefined) => (v?.endsWith('%') ? (num(v.slice(0, -1)) ?? 0) * 1000 : (num(v) ?? 0));
  const mods = children(el).map((m) => ({ name: m.name.slice(2), val: amount(m.attrs.val) }));
  return applyColorMods(base.toUpperCase(), mods);
}

const isColorElement = (c: XmlElement) => /^a:(srgbClr|schemeClr|sysClr|prstClr|scrgbClr|hslClr)$/.test(c.name);

function firstColorChild(el: XmlElement | undefined): XmlElement | undefined {
  if (!el) return undefined;
  return children(el).find(isColorElement);
}

function parseFill(spPr: XmlElement | undefined, ctx: DrawingContext): Fill | null | undefined {
  if (!spPr) return undefined;
  if (child(spPr, 'a:noFill')) return null;
  const solid = child(spPr, 'a:solidFill');
  if (solid) {
    const c = parseColor(firstColorChild(solid), ctx);
    return c ? { color: c.color, alpha: c.alpha } : undefined;
  }
  const grad = child(spPr, 'a:gradFill');
  if (grad) return gradientFill(grad, ctx);
  return undefined;
}

function gradientFill(grad: XmlElement, ctx: DrawingContext, phClr?: string): Fill | undefined {
  const g = parseGradient(grad, (el) => parseColor(el, ctx, phClr));
  return g ? { color: g.mid.color, alpha: g.mid.alpha, ...(g.gradient.stops.length > 1 && { gradient: g.gradient }) } : undefined;
}

/**
 * A shape style's fill (a:fillRef): the theme's fill of that index in the reference's colour. Only
 * the theme's gradients are taken from it; any other fill is the reference's colour.
 */
function styleFill(style: XmlElement | undefined, ctx: DrawingContext): Fill | null {
  const ref = styleRefColor(style, 'a:fillRef', ctx);
  if (!ref || ref.idx <= 0 || !ref.color) return null;
  const themed = ref.idx < 1000 ? ctx.theme.fillStyles[ref.idx - 1] : undefined;
  const grad = themed?.name === 'a:gradFill' ? gradientFill(themed, ctx, ref.color.color) : undefined;
  return grad ?? { color: ref.color.color, alpha: ref.color.alpha };
}

function parseLine(spPr: XmlElement | undefined, ctx: DrawingContext): LineStyle | null | undefined {
  const ln = child(spPr, 'a:ln');
  if (!ln) return undefined;
  if (child(ln, 'a:noFill')) return null;
  const width = ln.attrs.w ? emu(Number(ln.attrs.w)) : 0.75;
  const solid = child(ln, 'a:solidFill');
  const c = solid ? parseColor(firstColorChild(solid), ctx) : undefined;
  const dash = child(ln, 'a:prstDash')?.attrs.val ?? 'solid';
  if (!c) return { color: null, width, dash, alpha: 1, ...{ unresolved: true } } as LineStyle;
  return { color: c.color, width, dash, alpha: c.alpha };
}

function parsePos(el: XmlElement | undefined): PositionSpec {
  if (!el) return { relativeFrom: 'column', offset: 0 };
  const spec: PositionSpec = { relativeFrom: el.attrs.relativeFrom ?? 'column' };
  const off = child(el, 'wp:posOffset');
  const align = child(el, 'wp:align');
  const pctH = child(el, 'wp14:pctPosHOffset') ?? child(el, 'wp14:pctPosVOffset');
  if (off) spec.offset = emu(Number(textOf(off)) || 0);
  else if (align) spec.align = textOf(align).trim();
  else if (pctH) spec.pct = (Number(textOf(pctH)) || 0) / 100000;
  else spec.offset = 0;
  return spec;
}

function textOf(el: XmlElement): string {
  return el.children.filter((c): c is string => typeof c === 'string').join('');
}

function parseAnchor(a: XmlElement): AnchorProps {
  const wrapEl = children(a).find((c) => c.name.startsWith('wp:wrap'));
  let wrap: AnchorProps['wrap'] = 'none';
  let wrapText: AnchorProps['wrapText'] = 'bothSides';
  let polygon: AnchorProps['polygon'];
  if (wrapEl) {
    const kind = wrapEl.name.slice(7);
    wrap = kind === 'Square' ? 'square' : kind === 'Tight' ? 'tight' : kind === 'Through' ? 'through' : kind === 'TopAndBottom' ? 'topAndBottom' : 'none';
    const wt = wrapEl.attrs.wrapText;
    if (wt === 'left' || wt === 'right' || wt === 'largest') wrapText = wt;
    // Wrap polygon points are in 21600ths of the object's extent.
    const polyEl = child(wrapEl, 'wp:wrapPolygon');
    const points = polyEl ? children(polyEl).filter((c) => c.name === 'wp:start' || c.name === 'wp:lineTo') : [];
    const outline = points.map((c): [number, number] => [Number(c.attrs.x) / 21600, Number(c.attrs.y) / 21600]);
    if ((wrap === 'tight' || wrap === 'through') && outline.length >= 3 && outline.every(([x, y]) => Number.isFinite(x) && Number.isFinite(y))) polygon = outline;
  }
  const d = (n: string) => emu(Number(a.attrs[n] ?? 0) || 0);
  const simple = a.attrs.simplePos === '1' || a.attrs.simplePos === 'true';
  // Word 2010+ writes percentage positions as an mc:AlternateContent choice with a fallback.
  const alternates = children(a, 'mc:AlternateContent').flatMap(resolveAlternateContent);
  const position = (name: string) => child(a, name) ?? alternates.find((c) => c.name === name);
  let posH = parsePos(position('wp:positionH'));
  let posV = parsePos(position('wp:positionV'));
  if (simple) {
    const sp = child(a, 'wp:simplePos');
    posH = { relativeFrom: 'page', offset: emu(Number(sp?.attrs.x ?? 0)) };
    posV = { relativeFrom: 'page', offset: emu(Number(sp?.attrs.y ?? 0)) };
  }
  return {
    behindDoc: a.attrs.behindDoc === '1' || a.attrs.behindDoc === 'true',
    wrap,
    wrapText,
    distT: d('distT'),
    distB: d('distB'),
    distL: d('distL'),
    distR: d('distR'),
    ...(polygon && { polygon }),
    posH,
    posV,
    layoutInCell: a.attrs.layoutInCell !== '0' && a.attrs.layoutInCell !== 'false',
    allowOverlap: a.attrs.allowOverlap !== '0',
    zIndex: Number(a.attrs.relativeHeight ?? 0) || 0,
  };
}

function baseDrawing(kind: Drawing['kind'], inline: boolean, width: number, height: number): Drawing {
  return { kind, inline, width, height, effectL: 0, effectT: 0, effectR: 0, effectB: 0, rotation: 0, flipH: false, flipV: false, geometry: 'rect', hidden: false };
}

function applyXfrm(d: Drawing, spPr: XmlElement | undefined): void {
  const xfrm = child(spPr, 'a:xfrm');
  if (!xfrm) return;
  d.rotation = (Number(xfrm.attrs.rot ?? 0) || 0) / 60000;
  d.flipH = xfrm.attrs.flipH === '1';
  d.flipV = xfrm.attrs.flipV === '1';
}

function applyGeometry(d: Drawing, spPr: XmlElement | undefined): void {
  const prst = child(spPr, 'a:prstGeom');
  if (prst) {
    d.geometry = prst.attrs.prst ?? 'rect';
    const av = child(prst, 'a:avLst');
    for (const gd of children(av, 'a:gd')) {
      const m = /val\s+(-?\d+)/.exec(gd.attrs.fmla ?? '');
      if (m) (d.adjust ??= {})[gd.attrs.name ?? ''] = Number(m[1]);
    }
    return;
  }
  const cust = child(spPr, 'a:custGeom');
  if (cust) {
    d.geometry = 'custom';
    d.path = [];
    // A path without its own size is drawn in the shape's extent (EMU), as SmartArt connectors are.
    const ext = child(child(spPr, 'a:xfrm'), 'a:ext');
    for (const p of children(child(cust, 'a:pathLst'), 'a:path')) {
      const w = Number(p.attrs.w ?? 0) || Number(ext?.attrs.cx ?? 0) || 1;
      const h = Number(p.attrs.h ?? 0) || Number(ext?.attrs.cy ?? 0) || 1;
      const segs: Array<{ op: 'M' | 'L' | 'C' | 'Z'; pts: number[] }> = [];
      const pts = (el: XmlElement) => children(el, 'a:pt').flatMap((pt) => [(Number(pt.attrs.x) || 0) / w, (Number(pt.attrs.y) || 0) / h]);
      // The current point and the subpath's start: a quadratic curve begins at the current point.
      let [cx, cy, sx, sy] = [0, 0, 0, 0];
      for (const cmd of children(p)) {
        const q = pts(cmd);
        if (cmd.name === 'a:moveTo' && q.length >= 2) {
          segs.push({ op: 'M', pts: q.slice(0, 2) });
          [cx, cy, sx, sy] = [q[0], q[1], q[0], q[1]];
        } else if (cmd.name === 'a:lnTo' && q.length >= 2) {
          segs.push({ op: 'L', pts: q.slice(0, 2) });
          [cx, cy] = [q[0], q[1]];
        } else if (cmd.name === 'a:cubicBezTo' && q.length >= 6) {
          segs.push({ op: 'C', pts: q.slice(0, 6) });
          [cx, cy] = [q[4], q[5]];
        } else if (cmd.name === 'a:quadBezTo' && q.length >= 4) {
          // The same curve as a cubic: its control points lie two thirds of the way from each end to the quadratic's.
          const k = 2 / 3;
          segs.push({ op: 'C', pts: [cx + k * (q[0] - cx), cy + k * (q[1] - cy), q[2] + k * (q[0] - q[2]), q[3] + k * (q[1] - q[3]), q[2], q[3]] });
          [cx, cy] = [q[2], q[3]];
        } else if (cmd.name === 'a:arcTo') {
          const n = (a: string) => Number(cmd.attrs[a]) || 0;
          const [wR, hR, swAng] = [Math.abs(n('wR')), Math.abs(n('hR')), n('swAng')];
          if (wR && hR && swAng) {
            // The arc continues from the current point. Its angles are measured in path units, so it is
            // traced there and then scaled to the path like a point.
            const a = arc(0, 0, wR, hR, n('stAng') / 60000, swAng / 60000);
            const [ox, oy] = [cx * w - a.from[0], cy * h - a.from[1]];
            for (const c of a.segs) segs.push({ op: 'C', pts: c.pts.map((v, i) => (i % 2 === 0 ? (v + ox) / w : (v + oy) / h)) });
            const end = segs[segs.length - 1].pts;
            [cx, cy] = [end[4], end[5]];
          }
        } else if (cmd.name === 'a:close') {
          segs.push({ op: 'Z', pts: [] });
          [cx, cy] = [sx, sy];
        }
      }
      // A shape can fill one subpath and only outline another (an open arc beside a closed body).
      d.path.push({ fill: p.attrs.fill !== 'none', stroke: p.attrs.stroke !== '0' && p.attrs.stroke !== 'false', segs });
    }
  }
}

function styleRefColor(style: XmlElement | undefined, refName: string, ctx: DrawingContext): { idx: number; color?: { color: string; alpha: number } } | undefined {
  const ref = child(style, refName);
  if (!ref) return undefined;
  return { idx: Number(ref.attrs.idx ?? 0) || 0, color: parseColor(firstColorChild(ref), ctx) };
}

function parsePicture(pic: XmlElement, ctx: DrawingContext, d: Drawing): void {
  const blipFill = child(pic, 'pic:blipFill');
  d.kind = 'image';
  d.image = loadBlip(ctx, child(blipFill, 'a:blip'));
  const src = child(blipFill, 'a:srcRect');
  if (src) {
    const f = (n: string) => (Number(src.attrs[n] ?? 0) || 0) / 100000;
    const crop = { l: f('l'), t: f('t'), r: f('r'), b: f('b') };
    if (crop.l || crop.t || crop.r || crop.b) d.crop = crop;
  }
  const spPr = child(pic, 'pic:spPr');
  applyXfrm(d, spPr);
  applyGeometry(d, spPr);
  const line = parseLine(spPr, ctx);
  if (line && line.color) d.line = line;
}

function parseShape(wsp: XmlElement, ctx: DrawingContext, d: Drawing): void {
  d.kind = 'shape';
  const spPr = child(wsp, 'wps:spPr');
  applyXfrm(d, spPr);
  applyGeometry(d, spPr);
  const style = child(wsp, 'wps:style');
  let fill = parseFill(spPr, ctx);
  if (fill === undefined) fill = styleFill(style, ctx);
  let line = parseLine(spPr, ctx);
  const lnRef = styleRefColor(style, 'a:lnRef', ctx);
  if (line === undefined || (line && line.color === null)) {
    if (lnRef && lnRef.idx > 0 && lnRef.color) {
      const width = line?.width ?? [0, 0.5, 1, 1.5][Math.min(lnRef.idx, 3)];
      line = { color: lnRef.color.color, width, dash: line?.dash ?? 'solid', alpha: lnRef.color.alpha };
    } else line = null;
  }
  d.fill = fill;
  d.line = line;
  // Picture fill in a shape renders as an image.
  const blipFill = child(spPr, 'a:blipFill');
  if (blipFill) {
    d.image = loadBlip(ctx, child(blipFill, 'a:blip'));
    if (d.image) d.kind = 'image';
  }
  const txbx = child(wsp, 'wps:txbx');
  const content = child(txbx, 'w:txbxContent');
  const bodyPr = child(wsp, 'wps:bodyPr');
  if (content) {
    const ins = (n: string, dflt: number) => (bodyPr?.attrs[n] !== undefined ? emu(Number(bodyPr.attrs[n])) : dflt);
    const anchor = bodyPr?.attrs.anchor;
    const blocks = ctx.parseBlocks(content);
    // Text a styled shape's runs leave automatic takes the shape style's font colour (white text on
    // a "Colored Fill" shape).
    const fontColor = styleRefColor(style, 'a:fontRef', ctx)?.color?.color;
    if (fontColor) colorAutomaticText(blocks, fontColor);
    d.textBox = {
      blocks,
      insetL: ins('lIns', 7.2),
      insetT: ins('tIns', 3.6),
      insetR: ins('rIns', 7.2),
      insetB: ins('bIns', 3.6),
      vAlign: anchor === 'ctr' ? 'center' : anchor === 'b' ? 'bottom' : 'top',
      autoFit: !!child(bodyPr, 'a:spAutoFit'),
      vertical: bodyPr?.attrs.vert,
    };
  }
}

function colorAutomaticText(blocks: Block[], color: string): void {
  for (const b of blocks) {
    if (b.type === 'table') {
      for (const row of b.rows) for (const cell of row.cells) colorAutomaticText(cell.content, color);
      continue;
    }
    if (b.markProps.color === null) b.markProps = { ...b.markProps, color };
    for (const item of b.content) if ('props' in item && item.props.color === null) item.props = { ...item.props, color };
  }
}

function parseGroup(grp: XmlElement, ctx: DrawingContext, d: Drawing, isTop: boolean): void {
  d.kind = 'group';
  const grpSpPr = child(grp, 'wpg:grpSpPr') ?? child(grp, 'a:grpSpPr');
  const xfrm = child(grpSpPr, 'a:xfrm');
  const off = child(xfrm, 'a:off');
  const ext = child(xfrm, 'a:ext');
  const chOff = child(xfrm, 'a:chOff');
  const chExt = child(xfrm, 'a:chExt');
  const ox = Number(chOff?.attrs.x ?? off?.attrs.x ?? 0) || 0;
  const oy = Number(chOff?.attrs.y ?? off?.attrs.y ?? 0) || 0;
  const ecx = Number(ext?.attrs.cx ?? 0) || 1;
  const ecy = Number(ext?.attrs.cy ?? 0) || 1;
  const ccx = Number(chExt?.attrs.cx ?? ext?.attrs.cx ?? 0) || ecx;
  const ccy = Number(chExt?.attrs.cy ?? ext?.attrs.cy ?? 0) || ecy;
  // The top-level group is scaled to the drawing extent; nested groups use their own extents.
  const targetW = isTop ? d.width : emu(ecx);
  const targetH = isTop ? d.height : emu(ecy);
  const sx = targetW / emu(ccx);
  const sy = targetH / emu(ccy);
  if (!isTop) applyXfrm(d, grpSpPr);
  d.children = [];
  for (let c of children(grp)) {
    if (c.name === 'mc:AlternateContent') {
      const alt = resolveAlternateContent(c)[0];
      if (!alt) continue;
      c = alt;
    }
    let spPr: XmlElement | undefined;
    if (c.name === 'wps:wsp') spPr = child(c, 'wps:spPr');
    else if (c.name === 'pic:pic') spPr = child(c, 'pic:spPr');
    else if (c.name === 'wpg:grpSp') spPr = child(c, 'wpg:grpSpPr');
    else continue;
    const cx = child(spPr, 'a:xfrm');
    const coff = child(cx, 'a:off');
    const cext = child(cx, 'a:ext');
    const w = emu(Number(cext?.attrs.cx ?? 0) || 0) * sx;
    const h = emu(Number(cext?.attrs.cy ?? 0) || 0) * sy;
    const kid = baseDrawing('shape', true, w, h);
    kid.x = emu((Number(coff?.attrs.x ?? 0) || 0) - ox) * sx;
    kid.y = emu((Number(coff?.attrs.y ?? 0) || 0) - oy) * sy;
    if (c.name === 'wps:wsp') parseShape(c, ctx, kid);
    else if (c.name === 'pic:pic') parsePicture(c, ctx, kid);
    else parseGroup(c, ctx, kid, false);
    d.children.push(kid);
  }
}

/** Parse a w:drawing element. */
export function parseDrawing(el: XmlElement, ctx: DrawingContext): Drawing | null {
  const container = child(el, 'wp:inline') ?? child(el, 'wp:anchor');
  if (!container) return null;
  const inline = container.name === 'wp:inline';
  const extent = child(container, 'wp:extent');
  const d = baseDrawing('unsupported', inline, emu(Number(extent?.attrs.cx ?? 0) || 0), emu(Number(extent?.attrs.cy ?? 0) || 0));
  const ee = child(container, 'wp:effectExtent');
  if (ee) {
    d.effectL = emu(Number(ee.attrs.l ?? 0) || 0);
    d.effectT = emu(Number(ee.attrs.t ?? 0) || 0);
    d.effectR = emu(Number(ee.attrs.r ?? 0) || 0);
    d.effectB = emu(Number(ee.attrs.b ?? 0) || 0);
  }
  const docPr = child(container, 'wp:docPr');
  d.name = docPr?.attrs.name;
  const alt = docPr?.attrs.descr || docPr?.attrs.title;
  if (alt) d.alt = alt;
  d.hidden = docPr?.attrs.hidden === '1' || docPr?.attrs.hidden === 'true';
  if (!inline) d.anchor = parseAnchor(container);
  const gd = path(container, 'a:graphic', 'a:graphicData');
  if (!gd) return d;
  let content = children(gd)[0];
  if (content?.name === 'mc:AlternateContent') content = resolveAlternateContent(content)[0];
  if (!content) return d;
  if (content.name === 'pic:pic') parsePicture(content, ctx, d);
  else if (content.name === 'wps:wsp') parseShape(content, ctx, d);
  else if (content.name === 'wpg:wgp') parseGroup(content, ctx, d, true);
  else if (content.name === 'wpc:wpc') {
    // Drawing canvas: treat like a group of its shapes.
    parseGroup(content, ctx, d, true);
    easeGradients(d);
  } else if (content.name === 'dgm:relIds') {
    const pre = parseSmartArtDrawing(content, ctx, d);
    if (!pre) d.kind = 'unsupported';
    easeGradients(d);
  } else if (content.name === 'c:chart') {
    // Word keeps no picture of a chart, only its data: draw it from the chart part's cached values.
    const rel = ctx.pkg.rels(ctx.part).get(content.attrs['r:id'] ?? '');
    if (!rel || rel.external) return d;
    if (!ctx.charts.has(rel.target)) {
      const root = ctx.pkg.xml(rel.target);
      const space = root?.name === 'c:chartSpace' ? root : descendant(root, 'c:chartSpace');
      ctx.charts.set(rel.target, space && parseChart(space, ctx.theme, (e) => parseColor(e, ctx), ctx.chartBudget));
    }
    const chart = ctx.charts.get(rel.target);
    if (chart) {
      d.kind = 'shape';
      d.chart = chart;
    }
  }
  return d;
}

/** Mark the gradients of a canvas's or SmartArt's shapes as Word draws them there (GradientFill.eased). */
function easeGradients(d: Drawing): void {
  if (d.fill?.gradient) d.fill.gradient.eased = true;
  for (const kid of d.children ?? []) easeGradients(kid);
}

/** SmartArt keeps a pre-rendered shape tree in a "diagramDrawing" part; render that. */
function parseSmartArtDrawing(relIds: XmlElement, ctx: DrawingContext, d: Drawing): boolean {
  const dataRel = ctx.pkg.rels(ctx.part).get(relIds.attrs['r:dm'] ?? '');
  if (!dataRel) return false;
  const dataXml = ctx.pkg.xml(dataRel.target);
  const ext = descendant(dataXml, 'dsp:dataModelExt');
  const drawingRid = ext?.attrs.relId;
  if (!drawingRid) return false;
  const rel = ctx.pkg.rels(dataRel.target).get(drawingRid) ?? ctx.pkg.rels(ctx.part).get(drawingRid);
  if (!rel) return false;
  const drawing = ctx.pkg.xml(rel.target);
  const spTree = descendant(drawing, 'dsp:spTree');
  if (!spTree) return false;
  const sub: DrawingContext = { ...ctx, part: rel.target };
  const textColors = smartArtTextColors(relIds, dataXml, ctx);
  d.kind = 'group';
  d.children = [];
  for (const sp of children(spTree, 'dsp:sp')) {
    const spPr = child(sp, 'dsp:spPr');
    const xfrm = child(spPr, 'a:xfrm');
    const off = child(xfrm, 'a:off');
    const ext2 = child(xfrm, 'a:ext');
    const kid = baseDrawing('shape', true, emu(Number(ext2?.attrs.cx ?? 0) || 0), emu(Number(ext2?.attrs.cy ?? 0) || 0));
    kid.x = emu(Number(off?.attrs.x ?? 0) || 0);
    kid.y = emu(Number(off?.attrs.y ?? 0) || 0);
    applyXfrm(kid, spPr);
    applyGeometry(kid, spPr);
    const style = child(sp, 'dsp:style');
    let fill = parseFill(spPr, sub);
    if (fill === undefined) fill = styleFill(style, sub);
    let line = parseLine(spPr, sub);
    if (line === undefined || (line && line.color === null)) {
      const ref = styleRefColor(style, 'a:lnRef', sub);
      line = ref && ref.idx > 0 && ref.color ? { color: ref.color.color, width: line?.width ?? 0.75, dash: 'solid', alpha: ref.color.alpha } : null;
    }
    kid.fill = fill;
    kid.line = line;
    d.children.push(kid);
    const txBody = child(sp, 'dsp:txBody');
    if (!txBody) continue;
    const color = textColors.get(sp.attrs.modelId ?? '');
    const txXfrm = child(sp, 'dsp:txXfrm');
    const toff = child(txXfrm, 'a:off');
    const text = child(txXfrm, 'a:ext');
    if (!toff || !text) {
      kid.textBox = drawingMLTextBox(txBody, style, sub, color);
      continue;
    }
    // The text has a rectangle of its own, already inside the shape's text area, turned by the
    // shape's rotation plus its own and never mirrored: a chevron stood on end keeps its label
    // upright, a box turned on its side reads up the page. Like any shape text it may overflow.
    const label = baseDrawing('shape', true, emu(Number(text.attrs.cx ?? 0) || 0), emu(Number(text.attrs.cy ?? 0) || 0));
    label.x = emu(Number(toff.attrs.x ?? 0) || 0);
    label.y = emu(Number(toff.attrs.y ?? 0) || 0);
    label.rotation = kid.rotation + (Number(txXfrm!.attrs.rot ?? 0) || 0) / 60000;
    label.fill = null;
    label.line = null;
    label.textBox = { ...drawingMLTextBox(txBody, style, sub, color), unclipped: true };
    d.children.push(label);
  }
  return true;
}

/**
 * Each SmartArt node's text colour, from the diagram's colour definition by the node's style label.
 * Word draws the text in that colour even where the cached drawing's font reference disagrees (dark
 * text in white boxes that the cache records as white).
 */
function smartArtTextColors(relIds: XmlElement, dataXml: XmlElement | undefined, ctx: DrawingContext): Map<string, { color: string; alpha: number }> {
  const out = new Map<string, { color: string; alpha: number }>();
  const rel = ctx.pkg.rels(ctx.part).get(relIds.attrs['r:cs'] ?? '');
  if (!rel || !dataXml) return out;
  const lists = new Map<string, XmlElement>();
  for (const lbl of descendants(ctx.pkg.xml(rel.target), 'dgm:styleLbl')) {
    const list = child(lbl, 'dgm:txFillClrLst');
    if (lbl.attrs.name && firstColorChild(list)) lists.set(lbl.attrs.name, list!);
  }
  for (const pt of descendants(dataXml, 'dgm:pt')) {
    const set = child(pt, 'dgm:prSet');
    const list = pt.attrs.type === 'pres' ? lists.get(set?.attrs.presStyleLbl ?? '') : undefined;
    if (!list || !pt.attrs.modelId) continue;
    const colors = children(list).filter(isColorElement);
    const i = Number(set!.attrs.presStyleIdx ?? 0) || 0;
    // "repeat" keeps the last colour for later nodes, "cycle" starts over; "span" blends, taken as repeat.
    const c = parseColor(colors[list.attrs.meth === 'cycle' ? i % colors.length : Math.min(i, colors.length - 1)], ctx);
    if (c) out.set(pt.attrs.modelId, c);
  }
  return out;
}

const elem = (name: string, attrs: Record<string, string> = {}, kids: Array<XmlElement | string> = []): XmlElement => ({ name, attrs, children: kids }) as XmlElement;

/** Convert DrawingML text (a:p/a:r) into WordprocessingML paragraphs. */
function drawingMLTextBox(txBody: XmlElement, style: XmlElement | undefined, ctx: DrawingContext, textColor?: { color: string; alpha: number }): TextBox {
  const bodyPr = child(txBody, 'a:bodyPr');
  const fontRef = child(style, 'dsp:fontRef') ?? child(style, 'a:fontRef');
  const fontColor = textColor ?? parseColor(firstColorChild(fontRef), ctx);
  const scheme = fontRef?.attrs.idx === 'major' || fontRef?.attrs.idx === 'minor' ? fontRef.attrs.idx : undefined;
  const autofit = child(bodyPr, 'a:normAutofit');
  const fontScale = (Number(autofit?.attrs.fontScale ?? 100000) || 100000) / 100000;
  const lineScale = 1 - (Number(autofit?.attrs.lnSpcReduction ?? 0) || 0) / 100000;
  const twips = (pt: number) => String(Math.round(pt * 20));
  // A run's typeface: a name, a theme slot (+mn-lt, +mj-ea), or the shape style's theme font.
  const fonts = (rPr: XmlElement | undefined): Record<string, string> => {
    const a: Record<string, string> = {};
    const slot = (tag: string, nameAttrs: string[], themeAttrs: string[], suffix: string) => {
      const tf = child(rPr, tag)?.attrs.typeface;
      const theme = tf === undefined ? scheme : tf.startsWith('+mj') ? 'major' : tf.startsWith('+mn') ? 'minor' : undefined;
      if (tf && !theme) for (const n of nameAttrs) a[n] = tf;
      else if (theme) for (const n of themeAttrs) a[n] = theme + suffix;
    };
    slot('a:latin', ['w:ascii', 'w:hAnsi'], ['w:asciiTheme', 'w:hAnsiTheme'], 'HAnsi');
    slot('a:ea', ['w:eastAsia'], ['w:eastAsiaTheme'], 'EastAsia');
    slot('a:cs', ['w:cs'], ['w:cstheme'], 'Bidi');
    return a;
  };
  const runProps = (rPr: XmlElement | undefined, size: number): XmlElement[] => {
    const half = String(Math.max(1, Math.round(size * 2)));
    const out = [elem('w:sz', { 'w:val': half }), elem('w:szCs', { 'w:val': half }), elem('w:rFonts', fonts(rPr))];
    if (rPr?.attrs.b === '1') out.push(elem('w:b'), elem('w:bCs'));
    if (rPr?.attrs.i === '1') out.push(elem('w:i'), elem('w:iCs'));
    if (rPr?.attrs.u && rPr.attrs.u !== 'none') out.push(elem('w:u', { 'w:val': 'single' }));
    if (rPr?.attrs.strike && rPr.attrs.strike !== 'noStrike') out.push(elem('w:strike'));
    const baseline = Number(rPr?.attrs.baseline ?? 0) || 0;
    if (baseline) out.push(elem('w:vertAlign', { 'w:val': baseline > 0 ? 'superscript' : 'subscript' }));
    const col = parseColor(firstColorChild(child(rPr, 'a:solidFill')), ctx) ?? fontColor;
    if (col) out.push(elem('w:color', { 'w:val': col.color }));
    return out;
  };
  const ps = children(txBody, 'a:p');
  const firstLast = bodyPr?.attrs.spcFirstLastPara === '1';
  const paras = ps.map((p, pi) => {
    const pPr = child(p, 'a:pPr');
    const endSize = ((Number(child(p, 'a:endParaRPr')?.attrs.sz) || 0) / 100) * fontScale;
    const items = children(p).filter((c) => c.name === 'a:r' || c.name === 'a:fld' || c.name === 'a:br');
    const sizeOf = (rPr: XmlElement | undefined) => ((Number(rPr?.attrs.sz) || 0) / 100) * fontScale || endSize || 18 * fontScale;
    let size = 0;
    const runs: XmlElement[] = [];
    for (const r of items) {
      const rPr = child(r, 'a:rPr');
      const s = sizeOf(rPr);
      size = Math.max(size, s);
      const content =
        r.name === 'a:br'
          ? elem('w:br')
          : elem('w:t', { 'xml:space': 'preserve' }, [
              children(r, 'a:t')
                .map((t) => t.children.join(''))
                .join(''),
            ]);
      runs.push(elem('w:r', {}, [elem('w:rPr', {}, runProps(rPr, s)), content]));
    }
    size ||= endSize || 18 * fontScale;
    // A bullet hangs in the indent, in the first run's look unless the bullet names its own font
    // or colour. Empty paragraphs show no bullet.
    const bullet = child(pPr, 'a:buChar')?.attrs.char;
    const first = items.find((r) => r.name !== 'a:br');
    if (bullet && first) {
      const props = runProps(child(first, 'a:rPr'), sizeOf(child(first, 'a:rPr')));
      const face = child(pPr, 'a:buFont')?.attrs.typeface;
      const bc = parseColor(firstColorChild(child(pPr, 'a:buClr')), ctx);
      const adjusted = props.map((e) =>
        e.name === 'w:rFonts' && face && !face.startsWith('+') ? elem('w:rFonts', { 'w:ascii': face, 'w:hAnsi': face }) : e.name === 'w:color' && bc ? elem('w:color', { 'w:val': bc.color }) : e,
      );
      runs.unshift(elem('w:r', {}, [elem('w:rPr', {}, adjusted), elem('w:t', {}, [bullet]), elem('w:tab')]));
    }
    // Paragraph spacing: points, or a percentage of a single line (1.2 times the text size).
    const space = (tag: string) => {
      const s = child(pPr, tag);
      const pts = child(s, 'a:spcPts');
      if (pts) return (Number(pts.attrs.val) || 0) / 100;
      return ((Number(child(s, 'a:spcPct')?.attrs.val) || 0) / 100000) * size * 1.2;
    };
    const lnSpc = child(pPr, 'a:lnSpc');
    const lnPts = child(lnSpc, 'a:spcPts');
    const line = lnPts
      ? { 'w:line': twips(((Number(lnPts.attrs.val) || 0) / 100) * lineScale), 'w:lineRule': 'exact' }
      : { 'w:line': String(Math.round(240 * ((Number(child(lnSpc, 'a:spcPct')?.attrs.val ?? 100000) || 100000) / 100000) * lineScale)), 'w:lineRule': 'auto' };
    // Without spcFirstLastPara the first paragraph's space before and the last's space after are dropped.
    const before = pi === 0 && !firstLast ? 0 : space('a:spcBef');
    const after = pi === ps.length - 1 && !firstLast ? 0 : space('a:spcAft');
    const marL = emu(Number(pPr?.attrs.marL ?? 0) || 0);
    const indent = emu(Number(pPr?.attrs.indent ?? 0) || 0);
    const algn = pPr?.attrs.algn;
    const jc = algn === 'ctr' ? 'center' : algn === 'r' ? 'right' : algn === 'just' || algn === 'dist' ? 'both' : 'left';
    return elem('w:p', {}, [
      elem('w:pPr', {}, [
        elem('w:jc', { 'w:val': jc }),
        elem('w:spacing', { 'w:before': twips(before), 'w:after': twips(after), ...line }),
        elem('w:ind', { 'w:left': twips(marL), ...(indent < 0 ? { 'w:hanging': twips(-indent) } : { 'w:firstLine': twips(indent) }) }),
        elem('w:rPr', {}, runProps(child(p, 'a:endParaRPr'), size)),
      ]),
      ...runs,
    ]);
  });
  const ins = (n: string, dflt: number) => (bodyPr?.attrs[n] !== undefined ? emu(Number(bodyPr.attrs[n])) : dflt);
  const vert = bodyPr?.attrs.vert;
  const anchor = bodyPr?.attrs.anchor;
  return {
    blocks: ctx.parseBlocks(elem('w:txbxContent', {}, paras)),
    insetL: ins('lIns', 7.2),
    insetT: ins('tIns', 3.6),
    insetR: ins('rIns', 7.2),
    insetB: ins('bIns', 3.6),
    vAlign: anchor === 'ctr' ? 'center' : anchor === 'b' ? 'bottom' : 'top',
    autoFit: false,
    ...((vert === 'vert' || vert === 'vert270' || vert === 'eaVert') && { vertical: vert }),
  };
}

// ---------------------------------------------------------------------------------------------
// VML (legacy pictures, text boxes and watermarks)

function parseCss(style: string | undefined): Record<string, string> {
  const out: Record<string, string> = {};
  if (!style) return out;
  for (const part of style.split(';')) {
    const i = part.indexOf(':');
    if (i < 0) continue;
    out[part.slice(0, i).trim().toLowerCase()] = part.slice(i + 1).trim();
  }
  return out;
}

function cssLength(v: string | undefined): number | undefined {
  if (!v) return undefined;
  return measure(v, (n) => n * 0.75); // bare VML numbers are pixels
}

function vmlColor(v: string | undefined): Color | undefined {
  if (!v) return undefined;
  const c = v.split(/\s+/)[0];
  return hexColor(c) ?? undefined;
}

/** A VML fraction: a number, 16.16 fixed point ("32768f") or a percentage; `fallback` when absent or unreadable. */
function vmlFraction(v: string | undefined, fallback = 0): number {
  const n = !v ? NaN : v.endsWith('f') ? Number(v.slice(0, -1)) / 65536 : v.endsWith('%') ? Number(v.slice(0, -1)) / 100 : Number(v);
  return Number.isFinite(n) ? n : fallback;
}

// A VML shape names its shape type (type="#id") and takes the type's attributes, path and text
// area. Word writes each type once, before the first shape that uses it, so a document's types are
// collected as its pictures are read.
const vmlShapeTypes = new WeakMap<OpcPackage, Map<string, XmlElement>>();

// Shape types (o:spt) drawn as presets; the callouts' adjust values are the tail's tip.
const VML_PRESETS: Record<number, string> = {
  1: 'rect', 3: 'ellipse', 4: 'diamond', 6: 'rtTriangle', 20: 'line', 32: 'straightConnector1', 75: 'rect', 202: 'rect',
  61: 'wedgeRectCallout', 62: 'wedgeRoundRectCallout', 63: 'wedgeEllipseCallout',
};

/** Parse w:pict / w:object content. */
export function parseVml(el: XmlElement, ctx: DrawingContext): Drawing | null {
  let types = vmlShapeTypes.get(ctx.pkg);
  if (!types) vmlShapeTypes.set(ctx.pkg, (types = new Map()));
  for (const t of descendants(el, 'v:shapetype')) if (t.attrs.id) types.set(t.attrs.id, t);
  const shape = children(el).find((c) => /^v:(shape|rect|roundrect|oval|line|group|image)$/.test(c.name));
  if (!shape) return null;
  return parseVmlShape(shape, ctx, true);
}

/** The text area of a VML shape (its path's textboxrect, in the shape's coordinates) in points. */
function vmlTextRect(shape: XmlElement, type: XmlElement | undefined, width: number, height: number): [number, number, number, number] | undefined {
  const spec = child(shape, 'v:path')?.attrs.textboxrect ?? child(type, 'v:path')?.attrs.textboxrect;
  const r = spec?.split(',').map(Number);
  if (!r || r.length !== 4 || r.some((v) => !Number.isFinite(v))) return undefined;
  const [cw, ch] = (shape.attrs.coordsize ?? type?.attrs.coordsize ?? '21600,21600').split(',').map(Number);
  if (!(cw > 0 && ch > 0)) return undefined;
  return [(r[0] / cw) * width, (r[1] / ch) * height, width - (r[2] / cw) * width, height - (r[3] / ch) * height];
}

function parseVmlShape(shape: XmlElement, ctx: DrawingContext, top: boolean): Drawing | null {
  const css = parseCss(shape.attrs.style);
  const width = cssLength(css.width) ?? 0;
  const height = cssLength(css.height) ?? 0;
  const absolute = css.position === 'absolute';
  const d = baseDrawing('shape', !absolute || !top, width, height);
  if (shape.attrs.alt) d.alt = shape.attrs.alt;
  d.geometry = shape.name === 'v:oval' ? 'ellipse' : shape.name === 'v:roundrect' ? 'roundRect' : shape.name === 'v:line' ? 'line' : 'rect';
  const typeId = shape.name === 'v:shape' ? shape.attrs.type?.replace(/^#/, '') : undefined;
  const type = typeId ? vmlShapeTypes.get(ctx.pkg)?.get(typeId) : undefined;
  const attr = (name: string): string | undefined => shape.attrs[name] ?? type?.attrs[name];
  // Word's built-in types carry their number in their ID (_x0000_t63) even where the definition is missing.
  const spt = Number(type?.attrs['o:spt'] ?? /^_x0000_t(\d+)$/.exec(typeId ?? '')?.[1]);
  if (VML_PRESETS[spt]) {
    d.geometry = VML_PRESETS[spt];
    if (d.geometry.startsWith('wedge')) {
      const own = (shape.attrs.adj ?? '').split(',');
      const base = (type?.attrs.adj ?? '').split(',');
      const tip = (i: number, dflt: number) => Number(own[i]?.trim() || base[i]?.trim() || dflt);
      // VML puts the tip in the shape's 21600-unit square; DrawingML measures it from the centre.
      d.adjust = { adj1: ((tip(0, 1350) - 10800) / 21600) * 100000, adj2: ((tip(1, 25920) - 10800) / 21600) * 100000 };
    }
  }
  if (css.visibility === 'hidden') d.hidden = true;
  if (shape.attrs['o:hr'] === 't' && top && !absolute) {
    const align = shape.attrs['o:hralign'];
    // o:hrpct is in tenths of a percent; 0 keeps the shape's own width.
    d.hr = { align: align === 'center' || align === 'right' ? align : 'left', pct: (Number(shape.attrs['o:hrpct']) || 0) / 1000 };
  }
  if (css.rotation) d.rotation = parseFloat(css.rotation) || 0;
  if (css.flip) {
    d.flipH = css.flip.includes('x');
    d.flipV = css.flip.includes('y');
  }
  if (absolute && top) {
    const wrapEl = child(shape, 'w10:wrap');
    const wt = wrapEl?.attrs.type;
    const z = Number(css['z-index'] ?? 0) || 0;
    d.anchor = {
      behindDoc: z < 0,
      wrap: wt === 'square' ? 'square' : wt === 'tight' ? 'tight' : wt === 'through' ? 'through' : wt === 'topAndBottom' ? 'topAndBottom' : 'none',
      wrapText: 'bothSides',
      distT: cssLength(css['mso-wrap-distance-top']) ?? 0,
      distB: cssLength(css['mso-wrap-distance-bottom']) ?? 0,
      distL: cssLength(css['mso-wrap-distance-left']) ?? 9,
      distR: cssLength(css['mso-wrap-distance-right']) ?? 9,
      posH: {
        relativeFrom: vmlRelative(css['mso-position-horizontal-relative'], 'column'),
        offset: cssLength(css['margin-left']) ?? cssLength(css.left) ?? 0,
        align: css['mso-position-horizontal'] && css['mso-position-horizontal'] !== 'absolute' ? css['mso-position-horizontal'] : undefined,
      },
      posV: {
        relativeFrom: vmlRelative(css['mso-position-vertical-relative'], 'paragraph'),
        offset: cssLength(css['margin-top']) ?? cssLength(css.top) ?? 0,
        align: css['mso-position-vertical'] && css['mso-position-vertical'] !== 'absolute' ? css['mso-position-vertical'] : undefined,
      },
      layoutInCell: true,
      allowOverlap: true,
      zIndex: z,
    };
  } else if (!top) {
    d.x = cssLength(css.left) ?? 0;
    d.y = cssLength(css.top) ?? 0;
  }

  const off = (v: string | undefined) => v === 'f' || v === 'false';
  const fillEl = child(shape, 'v:fill');
  const filled = !off(attr('filled')) && !off(fillEl?.attrs.on);
  const opacity = Math.max(0, Math.min(1, vmlFraction(fillEl?.attrs.opacity, 1)));
  const fc = vmlColor(attr('fillcolor'));
  // A shape of a known type takes VML's defaults, a white fill and a black 0.75pt outline, unless
  // it or its type turns them off (pictures' type 75 does). Without the type, only what the shape
  // states is drawn.
  const typed = type !== undefined && spt !== 75;
  d.fill = filled && (attr('fillcolor') !== undefined || typed) ? { color: fc ?? 'FFFFFF', alpha: opacity } : null;
  const stroked = !off(attr('stroked')) && !off(child(shape, 'v:stroke')?.attrs.on);
  const outlined = attr('strokecolor') !== undefined || (shape.name !== 'v:shape' && attr('strokeweight') !== undefined) || typed;
  d.line = stroked && outlined
    ? { color: vmlColor(attr('strokecolor')) ?? '000000', width: cssLength(attr('strokeweight')) ?? 0.75, dash: 'solid', alpha: 1 }
    : null;

  const imagedata = child(shape, 'v:imagedata');
  if (imagedata) {
    d.kind = 'image';
    // VML's gain is the contrast slope and blacklevel half the brightness (Washout is gain 0.3,
    // blacklevel 0.35).
    const gain = vmlFraction(imagedata.attrs.gain, 1);
    d.image = adjusted(loadImage(ctx, imagedata.attrs['r:id'] ?? imagedata.attrs['r:pict']), Math.max(0, gain), 2 * vmlFraction(imagedata.attrs.blacklevel));
    const crop = {
      l: vmlFraction(imagedata.attrs.cropleft),
      t: vmlFraction(imagedata.attrs.croptop),
      r: vmlFraction(imagedata.attrs.cropright),
      b: vmlFraction(imagedata.attrs.cropbottom),
    };
    if (crop.l || crop.t || crop.r || crop.b) d.crop = crop;
    d.fill = null;
  }
  const textbox = child(shape, 'v:textbox');
  const content = child(textbox, 'w:txbxContent');
  if (content) {
    const inset = (textbox?.attrs.inset ?? '').split(',').map((s) => cssLength(s.trim()));
    const rect = vmlTextRect(shape, type, width, height);
    d.textBox = {
      blocks: ctx.parseBlocks(content),
      insetL: inset[0] ?? 7.2,
      insetT: inset[1] ?? 3.6,
      insetR: inset[2] ?? 7.2,
      insetB: inset[3] ?? 3.6,
      vAlign: 'top',
      autoFit: parseCss(textbox?.attrs.style)['mso-fit-shape-to-text'] === 't',
      // Word keeps the text of a turned legacy shape level, as Word 2003 drew it.
      upright: true,
      ...(rect && { rect }),
    };
  }
  const textpath = child(shape, 'v:textpath');
  if (textpath && textpath.attrs.string) {
    const tcss = parseCss(textpath.attrs.style);
    d.textPath = {
      text: textpath.attrs.string,
      font: (tcss['font-family'] ?? 'Calibri').replace(/^["']|["']$/g, ''),
      bold: tcss['font-weight'] === 'bold',
      italic: tcss['font-style'] === 'italic',
      color: fc ?? 'C0C0C0',
      alpha: opacity,
    };
    d.fill = null;
    d.line = null;
  }
  if (shape.name === 'v:group') {
    d.kind = 'group';
    const coordsize = (shape.attrs.coordsize ?? `${width},${height}`).split(',').map(Number);
    const coordorigin = (shape.attrs.coordorigin ?? '0,0').split(',').map(Number);
    const sx = width / (coordsize[0] || width || 1);
    const sy = height / (coordsize[1] || height || 1);
    d.children = [];
    for (const c of children(shape)) {
      if (!/^v:(shape|rect|roundrect|oval|line|group|image)$/.test(c.name)) continue;
      const k = parseVmlShape(c, ctx, false);
      if (!k) continue;
      // Child coordinates are in the group's coordinate space (unitless).
      const ccss = parseCss(c.attrs.style);
      k.x = ((Number(parseFloat(ccss.left ?? '0')) || 0) - (coordorigin[0] || 0)) * sx;
      k.y = ((Number(parseFloat(ccss.top ?? '0')) || 0) - (coordorigin[1] || 0)) * sy;
      k.width = (Number(parseFloat(ccss.width ?? '0')) || 0) * sx;
      k.height = (Number(parseFloat(ccss.height ?? '0')) || 0) * sy;
      d.children.push(k);
    }
  }
  if (d.kind === 'shape' && !d.fill && !d.line && !d.textBox && !d.textPath && !d.children) d.kind = 'unsupported';
  return d;
}

function vmlRelative(v: string | undefined, dflt: string): string {
  if (!v) return dflt;
  if (v === 'text') return dflt === 'paragraph' ? 'paragraph' : 'column';
  if (v === 'char') return 'character';
  return v;
}
