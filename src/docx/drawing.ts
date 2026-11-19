import { child, children, descendant, path, resolveAlternateContent, type XmlElement } from '../util/xml.js';
import type { AnchorProps, Block, Color, Drawing, Fill, ImageRef, LineStyle, PositionSpec } from './model.js';
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
}

const PRESET_COLORS: Record<string, string> = { black: '000000', white: 'FFFFFF', red: 'FF0000', green: '008000', blue: '0000FF', yellow: 'FFFF00', gray: '808080' };

export function loadImage(ctx: DrawingContext, rid: string | undefined): ImageRef | undefined {
  if (!rid) return undefined;
  const rel = ctx.pkg.rels(ctx.part).get(rid);
  if (!rel || rel.external) return undefined;
  const cached = ctx.images.get(rel.target);
  if (cached) return cached;
  const data = ctx.pkg.bytes(rel.target);
  if (!data) return undefined;
  const ref: ImageRef = { id: rel.target, data };
  ctx.images.set(rel.target, ref);
  return ref;
}

function loadBlip(ctx: DrawingContext, blip: XmlElement | undefined): ImageRef | undefined {
  const fallback = loadImage(ctx, blip?.attrs['r:embed']);
  const svgBlip = children(child(blip, 'a:extLst'), 'a:ext').map((ext) => child(ext, 'asvg:svgBlip')).find(Boolean);
  const svg = loadImage(ctx, svgBlip?.attrs['r:embed']);
  return svg && fallback ? { ...fallback, svg } : svg ?? fallback;
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
      const c = (v: string | undefined) => Math.round(((num(v) ?? 0) / 100000) * 255);
      base = [c(el.attrs.r), c(el.attrs.g), c(el.attrs.b)].map((v) => v.toString(16).padStart(2, '0')).join('');
      break;
    }
    default:
      return undefined;
  }
  if (!base) return undefined;
  const mods = children(el).map((m) => ({ name: m.name.slice(2), val: num(m.attrs.val) ?? 0 }));
  return applyColorMods(base.toUpperCase(), mods);
}

function firstColorChild(el: XmlElement | undefined): XmlElement | undefined {
  if (!el) return undefined;
  return children(el).find((c) => /^a:(srgbClr|schemeClr|sysClr|prstClr|scrgbClr|hslClr)$/.test(c.name));
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
  if (grad) {
    const stops = children(child(grad, 'a:gsLst'), 'a:gs');
    const mid = stops[Math.floor(stops.length / 2)] ?? stops[0];
    const c = parseColor(firstColorChild(mid), ctx);
    return c ? { color: c.color, alpha: c.alpha } : undefined;
  }
  return undefined;
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
  if (wrapEl) {
    const kind = wrapEl.name.slice(7);
    wrap = kind === 'Square' ? 'square' : kind === 'Tight' ? 'tight' : kind === 'Through' ? 'through' : kind === 'TopAndBottom' ? 'topAndBottom' : 'none';
    const wt = wrapEl.attrs.wrapText;
    if (wt === 'left' || wt === 'right' || wt === 'largest') wrapText = wt;
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
    for (const p of children(child(cust, 'a:pathLst'), 'a:path')) {
      const w = Number(p.attrs.w ?? 0) || 1;
      const h = Number(p.attrs.h ?? 0) || 1;
      const sub: Array<{ op: 'M' | 'L' | 'C' | 'Z'; pts: number[] }> = [];
      const pts = (el: XmlElement) => children(el, 'a:pt').flatMap((pt) => [(Number(pt.attrs.x) || 0) / w, (Number(pt.attrs.y) || 0) / h]);
      for (const cmd of children(p)) {
        if (cmd.name === 'a:moveTo') sub.push({ op: 'M', pts: pts(cmd) });
        else if (cmd.name === 'a:lnTo') sub.push({ op: 'L', pts: pts(cmd) });
        else if (cmd.name === 'a:cubicBezTo') sub.push({ op: 'C', pts: pts(cmd) });
        else if (cmd.name === 'a:quadBezTo') {
          const q = pts(cmd);
          sub.push({ op: 'C', pts: [q[0], q[1], q[0], q[1], q[2], q[3]] });
        } else if (cmd.name === 'a:close') sub.push({ op: 'Z', pts: [] });
      }
      if (p.attrs.fill === 'none') (sub as unknown as { noFill: boolean }).noFill = true;
      d.path.push(sub);
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
  if (fill === undefined) {
    const ref = styleRefColor(style, 'a:fillRef', ctx);
    fill = ref && ref.idx > 0 && ref.color ? { color: ref.color.color, alpha: ref.color.alpha } : null;
  }
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
    d.textBox = {
      blocks: ctx.parseBlocks(content),
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
  } else if (content.name === 'dgm:relIds') {
    const pre = parseSmartArtDrawing(content, ctx, d);
    if (!pre) d.kind = 'unsupported';
  }
  return d;
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
    if (fill === undefined) {
      const ref = styleRefColor(style, 'a:fillRef', sub);
      fill = ref && ref.idx > 0 && ref.color ? { color: ref.color.color, alpha: ref.color.alpha } : null;
    }
    let line = parseLine(spPr, sub);
    if (line === undefined || (line && line.color === null)) {
      const ref = styleRefColor(style, 'a:lnRef', sub);
      line = ref && ref.idx > 0 && ref.color ? { color: ref.color.color, width: line?.width ?? 0.75, dash: 'solid', alpha: ref.color.alpha } : null;
    }
    kid.fill = fill;
    kid.line = line;
    const txBody = child(sp, 'dsp:txBody');
    if (txBody) kid.textBox = drawingMLTextBox(txBody, child(sp, 'dsp:txXfrm'), kid, style, sub);
    d.children.push(kid);
  }
  return true;
}

/** Convert DrawingML text (a:p/a:r) into simple paragraphs via a synthetic WordprocessingML tree. */
function drawingMLTextBox(txBody: XmlElement, txXfrm: XmlElement | undefined, shape: Drawing, style: XmlElement | undefined, ctx: DrawingContext) {
  const bodyPr = child(txBody, 'a:bodyPr');
  const fontColor = parseColor(firstColorChild(child(style, 'dsp:fontRef') ?? child(style, 'a:fontRef')), ctx);
  const paras: XmlElement[] = [];
  for (const p of children(txBody, 'a:p')) {
    const pPr = child(p, 'a:pPr');
    const algn = pPr?.attrs.algn;
    const jc = algn === 'ctr' ? 'center' : algn === 'r' ? 'right' : algn === 'just' ? 'both' : 'left';
    const runs: XmlElement[] = [];
    for (const r of children(p, 'a:r')) {
      const rPr = child(r, 'a:rPr');
      const sz = Number(rPr?.attrs.sz ?? child(child(p, 'a:endParaRPr'), 'x')?.attrs.sz ?? 1800) / 100;
      const col = parseColor(firstColorChild(child(rPr, 'a:solidFill')), ctx) ?? fontColor;
      const props: XmlElement[] = [
        { name: 'w:sz', attrs: { 'w:val': String(Math.round(sz * 2)) }, children: [] },
        { name: 'w:szCs', attrs: { 'w:val': String(Math.round(sz * 2)) }, children: [] },
      ];
      if (rPr?.attrs.b === '1') props.push({ name: 'w:b', attrs: {}, children: [] });
      if (rPr?.attrs.i === '1') props.push({ name: 'w:i', attrs: {}, children: [] });
      if (col) props.push({ name: 'w:color', attrs: { 'w:val': col.color }, children: [] });
      const latin = child(rPr, 'a:latin')?.attrs.typeface;
      if (latin && !latin.startsWith('+')) props.push({ name: 'w:rFonts', attrs: { 'w:ascii': latin, 'w:hAnsi': latin }, children: [] });
      const text = children(r, 'a:t')
        .map((t) => t.children.join(''))
        .join('');
      runs.push({
        name: 'w:r',
        attrs: {},
        children: [
          { name: 'w:rPr', attrs: {}, children: props },
          { name: 'w:t', attrs: { 'xml:space': 'preserve' }, children: [text] },
        ],
      });
    }
    paras.push({
      name: 'w:p',
      attrs: {},
      children: [
        {
          name: 'w:pPr',
          attrs: {},
          children: [
            { name: 'w:jc', attrs: { 'w:val': jc }, children: [] },
            { name: 'w:spacing', attrs: { 'w:before': '0', 'w:after': '0', 'w:line': '216', 'w:lineRule': 'auto' }, children: [] },
          ],
        },
        ...runs,
      ],
    });
  }
  const container: XmlElement = { name: 'w:txbxContent', attrs: {}, children: paras };
  const ins = (n: string, dflt: number) => (bodyPr?.attrs[n] !== undefined ? emu(Number(bodyPr.attrs[n])) : dflt);
  // txXfrm gives the text rectangle inside the shape.
  if (txXfrm) {
    const off = child(txXfrm, 'a:off');
    const ext = child(txXfrm, 'a:ext');
    if (off && ext && shape.x !== undefined && shape.y !== undefined) {
      const tx = emu(Number(off.attrs.x) || 0) - shape.x;
      const ty = emu(Number(off.attrs.y) || 0) - shape.y;
      const tw = emu(Number(ext.attrs.cx) || 0);
      const th = emu(Number(ext.attrs.cy) || 0);
      return {
        blocks: ctx.parseBlocks(container),
        insetL: tx + ins('lIns', 7.2),
        insetT: ty + ins('tIns', 3.6),
        insetR: shape.width - tx - tw + ins('rIns', 7.2),
        insetB: shape.height - ty - th + ins('bIns', 3.6),
        vAlign: (bodyPr?.attrs.anchor === 't' ? 'top' : bodyPr?.attrs.anchor === 'b' ? 'bottom' : 'center') as 'top' | 'center' | 'bottom',
        autoFit: false,
      };
    }
  }
  return {
    blocks: ctx.parseBlocks(container),
    insetL: ins('lIns', 7.2),
    insetT: ins('tIns', 3.6),
    insetR: ins('rIns', 7.2),
    insetB: ins('bIns', 3.6),
    vAlign: 'center' as const,
    autoFit: false,
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

function vmlFraction(v: string | undefined): number {
  if (!v) return 0;
  if (v.endsWith('f')) return (Number(v.slice(0, -1)) || 0) / 65536;
  return Number(v) || 0;
}

/** Parse w:pict / w:object content. */
export function parseVml(el: XmlElement, ctx: DrawingContext): Drawing | null {
  const shape = children(el).find((c) => /^v:(shape|rect|roundrect|oval|line|group|image)$/.test(c.name));
  if (!shape) return null;
  return parseVmlShape(shape, ctx, true);
}

function parseVmlShape(shape: XmlElement, ctx: DrawingContext, top: boolean): Drawing | null {
  const css = parseCss(shape.attrs.style);
  const width = cssLength(css.width) ?? 0;
  const height = cssLength(css.height) ?? 0;
  const absolute = css.position === 'absolute';
  const d = baseDrawing('shape', !absolute || !top, width, height);
  d.geometry = shape.name === 'v:oval' ? 'ellipse' : shape.name === 'v:roundrect' ? 'roundRect' : shape.name === 'v:line' ? 'line' : 'rect';
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

  const filled = shape.attrs.filled !== 'f' && shape.attrs.filled !== 'false';
  const fillEl = child(shape, 'v:fill');
  const opacity = fillEl?.attrs.opacity ? vmlFraction(fillEl.attrs.opacity) || 1 : 1;
  const fc = vmlColor(shape.attrs.fillcolor);
  d.fill = filled && shape.attrs.fillcolor !== undefined ? { color: fc ?? 'FFFFFF', alpha: opacity } : null;
  const stroked = shape.attrs.stroked !== 'f' && shape.attrs.stroked !== 'false';
  d.line = stroked && (shape.attrs.strokecolor !== undefined || shape.attrs.strokeweight !== undefined) && shape.name !== 'v:shape'
    ? { color: vmlColor(shape.attrs.strokecolor) ?? '000000', width: cssLength(shape.attrs.strokeweight) ?? 0.75, dash: 'solid', alpha: 1 }
    : null;
  if (stroked && shape.name === 'v:shape' && shape.attrs.strokecolor) {
    d.line = { color: vmlColor(shape.attrs.strokecolor) ?? '000000', width: cssLength(shape.attrs.strokeweight) ?? 0.75, dash: 'solid', alpha: 1 };
  }

  const imagedata = child(shape, 'v:imagedata');
  if (imagedata) {
    d.kind = 'image';
    d.image = loadImage(ctx, imagedata.attrs['r:id'] ?? imagedata.attrs['r:pict']);
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
    d.textBox = {
      blocks: ctx.parseBlocks(content),
      insetL: inset[0] ?? 7.2,
      insetT: inset[1] ?? 3.6,
      insetR: inset[2] ?? 7.2,
      insetB: inset[3] ?? 3.6,
      vAlign: 'top',
      autoFit: parseCss(textbox?.attrs.style)['mso-fit-shape-to-text'] === 't',
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
