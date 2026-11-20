import svgPath from 'svgpath';
import { children, parseXml, type XmlElement } from '../util/xml.js';

const PRESENTATION = new Set(['fill', 'fill-rule', 'stroke', 'stroke-width', 'stroke-linecap', 'stroke-linejoin',
  'stroke-miterlimit', 'stroke-dasharray', 'stroke-dashoffset', 'color', 'display']);
const METADATA = new Set(['metadata', 'title', 'desc', 'defs']);
const GEOMETRY: Record<string, string[]> = {
  svg: ['width', 'height', 'viewBox', 'preserveAspectRatio', 'version', 'baseProfile'],
  g: [],
  path: ['d'],
  rect: ['x', 'y', 'width', 'height', 'rx', 'ry'],
  circle: ['cx', 'cy', 'r'],
  ellipse: ['cx', 'cy', 'rx', 'ry'],
  line: ['x1', 'y1', 'x2', 'y2'],
  polygon: ['points'],
  polyline: ['points'],
};
const COLORS: Record<string, string> = {
  black: '000000', silver: 'c0c0c0', gray: '808080', white: 'ffffff', maroon: '800000', red: 'ff0000',
  purple: '800080', fuchsia: 'ff00ff', green: '008000', lime: '00ff00', olive: '808000', yellow: 'ffff00',
  navy: '000080', blue: '0000ff', teal: '008080', aqua: '00ffff',
};
const NUMBER = /^[+-]?(?:\d+\.?\d*|\.\d+)(?:e[+-]?\d+)?$/i;
const unsupported = Symbol('unsupported SVG');

function numeric(value: string): number {
  if (!NUMBER.test(value)) throw unsupported;
  const n = Number(value);
  if (!Number.isFinite(n) || Math.abs(n) > 1e8) throw unsupported;
  return n;
}

function length(value = '0'): number {
  return numeric(value.trim().replace(/px$/, ''));
}

function numbers(value: string): number[] {
  return value.trim().split(/\s*,\s*|\s+/).map(numeric);
}

function pdfNum(n: number): string {
  if (!Number.isFinite(n) || Math.abs(n) > 1e10) throw unsupported;
  return Number(n.toFixed(10)) === 0 ? '0' : n.toFixed(10).replace(/\.?0+$/, '');
}

function operands(values: number[]): string {
  return values.map(pdfNum).join(' ');
}

function viewport(attrs: XmlElement['attrs']): string {
  const box = attrs.viewBox ? numbers(attrs.viewBox) : undefined;
  if (box && (box.length !== 4 || box[2] <= 0 || box[3] <= 0)) throw unsupported;
  const width = attrs.width && attrs.width !== '100%' ? length(attrs.width) : box?.[2];
  const height = attrs.height && attrs.height !== '100%' ? length(attrs.height) : box?.[3];
  if (!width || !height || width < 0 || height < 0) throw unsupported;
  const [x, y, w, h] = box ?? [0, 0, width, height];
  let sx = width / w;
  let sy = height / h;
  let dx = 0;
  let dy = 0;
  const align = attrs.preserveAspectRatio?.trim() ?? 'xMidYMid';
  if (box && align !== 'none') {
    const match = /^x(Min|Mid|Max)Y(Min|Mid|Max)(?:\s+(meet|slice))?$/.exec(align);
    if (!match) throw unsupported;
    sx = sy = match[3] === 'slice' ? Math.max(sx, sy) : Math.min(sx, sy);
    const fraction = { Min: 0, Mid: 0.5, Max: 1 };
    dx = (width - w * sx) * fraction[match[1] as keyof typeof fraction];
    dy = (height - h * sy) * fraction[match[2] as keyof typeof fraction];
  }
  return operands([sx / width, 0, 0, -sy / height, (dx - x * sx) / width, 1 - (dy - y * sy) / height]) + ' cm';
}

function transform(value: string): string {
  const expected = { matrix: [6], translate: [1, 2], scale: [1, 2], rotate: [1, 3], skewX: [1], skewY: [1] };
  const pattern = /(?:\s*,?\s*)(matrix|translate|scale|rotate|skewX|skewY)\s*\(([^()]*)\)/y;
  const operators: string[] = [];
  value = value.trim();
  for (let end = 0; end < value.length; ) {
    const match = pattern.exec(value);
    if (!match || operators.length >= 32) throw unsupported;
    end = pattern.lastIndex;
    const args = numbers(match[2]);
    if (!expected[match[1] as keyof typeof expected].includes(args.length)) throw unsupported;
    operators.push(`${match[1]}(${args.join(' ')})`);
  }
  const points: number[][] = [];
  svgPath('M0 0L1 0L0 1').transform(operators.join(' ')).iterate(([, ...point]) => { points.push(point); });
  const [[x, y], [xx, xy], [yx, yy]] = points;
  return operands([xx - x, xy - y, yx - x, yy - y, x, y]) + ' cm';
}

function color(value: string, current: string, stroke: boolean): string {
  value = value.trim().toLowerCase();
  if (value === 'currentcolor') value = current.trim().toLowerCase();
  if (value === 'none') return '';
  let rgb: number[];
  const hex = Object.hasOwn(COLORS, value) ? COLORS[value] : (/^#(?:[0-9a-f]{3}|[0-9a-f]{6})$/.test(value) ? value.slice(1) : undefined);
  if (hex) {
    const pairs = hex.length === 3 ? [...hex].map((ch) => ch + ch) : hex.match(/../g)!;
    rgb = pairs.map((pair) => parseInt(pair, 16) / 255);
  } else {
    const match = /^rgb\(([^()]+)\)$/.exec(value);
    if (!match) throw unsupported;
    rgb = match[1].trim().split(/\s*,\s*|\s+/).map((v) => {
      const scaled = v.endsWith('%') ? numeric(v.slice(0, -1)) / 100 : numeric(v) / 255;
      return Math.max(0, Math.min(1, scaled));
    });
    if (rgb.length !== 3) throw unsupported;
  }
  return operands(rgb) + (stroke ? ' RG' : ' rg');
}

function strokeStyle(style: XmlElement['attrs'], width: number): string {
  const cap = ['butt', 'round', 'square'].indexOf(style['stroke-linecap'] ?? 'butt');
  const join = ['miter', 'round', 'bevel'].indexOf(style['stroke-linejoin'] ?? 'miter');
  const limit = length(style['stroke-miterlimit'] ?? '4');
  const dash = style['stroke-dasharray'] && style['stroke-dasharray'] !== 'none' ? numbers(style['stroke-dasharray']) : [];
  if (cap < 0 || join < 0 || limit < 1 || dash.some((n) => n < 0)) throw unsupported;
  if (dash.length % 2) dash.push(...dash);
  const total = dash.reduce((sum, n) => sum + n, 0);
  const offset = length(style['stroke-dashoffset']);
  const pattern = total > 0 ? `[${operands(dash)}] ${pdfNum(((offset % total) + total) % total)}` : '[] 0';
  return `${pdfNum(width)} w ${cap} J ${join} j ${pdfNum(limit)} M ${pattern} d`;
}

function geometry(tag: string, a: XmlElement['attrs']): string {
  switch (tag) {
    case 'path':
      return a.d ?? '';
    case 'line':
      return `M ${length(a.x1)} ${length(a.y1)} L ${length(a.x2)} ${length(a.y2)}`;
    case 'circle':
    case 'ellipse': {
      const cx = length(a.cx), cy = length(a.cy);
      const rx = length(a.r ?? a.rx), ry = length(a.r ?? a.ry);
      return rx > 0 && ry > 0 ? `M ${cx - rx} ${cy} a ${rx} ${ry} 0 1 0 ${2 * rx} 0 a ${rx} ${ry} 0 1 0 ${-2 * rx} 0 Z` : '';
    }
    case 'rect': {
      const x = length(a.x), y = length(a.y), w = length(a.width), h = length(a.height);
      const rx = Math.min(w / 2, length(a.rx ?? a.ry)), ry = Math.min(h / 2, length(a.ry ?? a.rx));
      if (w <= 0 || h <= 0) return '';
      if (rx <= 0 || ry <= 0) return `M ${x} ${y} h ${w} v ${h} h ${-w} Z`;
      const arc = `a ${rx} ${ry} 0 0 1`;
      return `M ${x + rx} ${y} h ${w - 2 * rx} ${arc} ${rx} ${ry} v ${h - 2 * ry} ${arc} ${-rx} ${ry} h ${2 * rx - w} ${arc} ${-rx} ${-ry} v ${2 * ry - h} ${arc} ${rx} ${-ry} Z`;
    }
    case 'polygon':
    case 'polyline': {
      if (!a.points?.trim()) return '';
      const points = numbers(a.points);
      if (points.length % 2) throw unsupported;
      return `M ${points.join(' ')}${tag === 'polygon' ? ' Z' : ''}`;
    }
    default:
      return '';
  }
}

function pdfPath(path: string): string {
  const parsed = svgPath(path);
  if ('err' in parsed && parsed.err) throw unsupported;
  const out: string[] = [];
  parsed.abs().unshort().unarc().iterate((segment, _, x, y) => {
    const cmd = segment[0];
    if (cmd === 'M' || cmd === 'L' || cmd === 'C') out.push(operands(segment.slice(1) as number[]) + ' ' + cmd.toLowerCase());
    if (cmd === 'H') out.push(operands([segment[1], y]) + ' l');
    if (cmd === 'V') out.push(operands([x, segment[1]]) + ' l');
    if (cmd === 'Q') {
      const [, qx, qy, ex, ey] = segment;
      out.push(operands([x + (qx - x) * 2 / 3, y + (qy - y) * 2 / 3,
        ex + (qx - ex) * 2 / 3, ey + (qy - ey) * 2 / 3, ex, ey]) + ' c');
    }
    if (cmd === 'Z' || cmd === 'z') out.push('h');
  });
  return out.join('\n');
}

/** Compile a static SVG into a unit-square PDF form using only drawing operators. */
export function svgToPdfForm(data: Uint8Array): string | null {
  // A compact, bounded vector subset. Unsupported visible content uses Word's preview if present.
  if (data.length > 512 * 1024) return null;
  try {
    const source = new TextDecoder('utf-8', { fatal: true }).decode(data);
    if (/<\?xml-stylesheet\b|<(?:[\w.-]+:)?style[\s/>]/.test(source)) return null;
    const root = parseXml(source);
    if (root.name !== 'svg:svg' && root.name !== 'svg') return null;
    const out = [viewport(root.attrs)];
    const walk = (el: XmlElement, inherited: XmlElement['attrs'], depth: number) => {
      const tag = el.name.replace(/^svg:/, '');
      if (METADATA.has(tag)) return;
      if (depth > 32 || !Object.hasOwn(GEOMETRY, tag) || (depth > 0 && tag === 'svg')) throw unsupported;
      const a = el.attrs;
      const style = { ...inherited };
      delete style.display;
      for (const [key, value] of Object.entries(a)) {
        if (PRESENTATION.has(key)) style[key] = value.trim();
        else if (!GEOMETRY[tag].includes(key) && !['id', 'class', 'style', 'transform', 'role', 'xml:space'].includes(key) &&
          !/^(?:aria-|data-|on|inkscape:|sodipodi:)/.test(key)) throw unsupported;
      }
      for (const declaration of (a.style ?? '').split(';')) {
        if (!declaration.trim()) continue;
        const match = /^\s*([a-z-]+)\s*:\s*(.*?)\s*$/.exec(declaration);
        if (!match || !PRESENTATION.has(match[1])) throw unsupported;
        style[match[1]] = match[2];
      }
      if (style.display === 'none') return;
      if (a.transform) out.push('q', transform(a.transform));
      const shape = geometry(tag, a);
      if (shape) {
        const current = style.color ?? 'black';
        const fill = tag !== 'line' ? color(style.fill ?? 'black', current, false) : '';
        const width = length(style['stroke-width'] ?? '1');
        const stroke = width > 0 ? color(style.stroke ?? 'none', current, true) : '';
        const rule = style['fill-rule'] ?? 'nonzero';
        if (rule !== 'nonzero' && rule !== 'evenodd') throw unsupported;
        if (fill || stroke) {
          if (fill) out.push(fill);
          if (stroke) out.push(stroke, strokeStyle(style, width));
          out.push(pdfPath(shape), (fill ? (stroke ? 'B' : 'f') + (rule === 'evenodd' ? '*' : '') : 'S'));
        }
      }
      for (const child of children(el)) walk(child, style, depth + 1);
      if (a.transform) out.push('Q');
    };
    walk(root, {}, 0);
    return out.join('\n');
  } catch (error) {
    if (error === unsupported || error instanceof Error) return null;
    throw error;
  }
}
