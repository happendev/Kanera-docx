// Colour glyphs as PDF form XObjects: each glyph's paint tree (see fonts/colr.ts) drawn in font
// units, with outlines as clipping paths, gradients as shadings, alpha as soft masks and COLR's
// composite modes as soft masks or blend modes. A form is written once per glyph and text colour.

import { BLEND_MODES, COMPOSITE, colorGlyph, glyphOutline, multiply, type ColorPaint, type ColorStop, type Matrix, type PaintColor } from '../fonts/colr.js';
import type { SfntFont } from '../fonts/sfnt.js';
import { latin1, num, type PdfObjects } from './objects.js';

type Box = [number, number, number, number];

/** A form's resources, named as the form's content uses them. */
class Resources {
  private readonly states = new Map<string, string>();
  private readonly shadings: string[] = [];
  private readonly forms: string[] = [];

  state(dict: string): string {
    let name = this.states.get(dict);
    if (!name) this.states.set(dict, (name = `G${this.states.size + 1}`));
    return name;
  }

  shading(dict: string): string {
    const name = `S${this.shadings.length + 1}`;
    this.shadings.push(`/${name} ${dict}`);
    return name;
  }

  form(obj: number): string {
    const name = `X${this.forms.length + 1}`;
    this.forms.push(`/${name} ${obj} 0 R`);
    return name;
  }

  toString(): string {
    const states = [...this.states].map(([dict, name]) => `/${name} ${dict}`).join(' ');
    return `<<${states ? ` /ExtGState << ${states} >>` : ''}${this.shadings.length ? ` /Shading << ${this.shadings.join(' ')} >>` : ''}${this.forms.length ? ` /XObject << ${this.forms.join(' ')} >>` : ''} >>`;
  }
}

const INVERT = '<< /FunctionType 2 /Domain [0 1] /C0 [1] /C1 [0] /N 1 >>';

export class ColorGlyphForms {
  private readonly forms = new Map<SfntFont, Map<string, number | null>>();

  constructor(
    private readonly pdf: PdfObjects,
    private readonly compress: boolean,
  ) {}

  /**
   * The form drawing glyph `gid` of `font` in font units, or null when it has no colour glyph. A
   * glyph without one draws its outline in the text colour when `plain` is set. `group` makes the
   * form a transparency group, so translucent text fades the glyph as a whole, not layer by layer.
   */
  form(font: SfntFont, gid: number, textColor: string, opts: { plain?: boolean; group?: boolean } = {}): number | null {
    let byKey = this.forms.get(font);
    if (!byKey) this.forms.set(font, (byKey = new Map()));
    const glyph = colorGlyph(font, gid);
    if (!glyph && !opts.plain) return null;
    const fg = textColor.slice(0, 6).toUpperCase();
    const key = `${gid}:${!glyph || glyph.foreground ? fg : ''}:${opts.group ? 'g' : ''}`;
    let obj = byKey.get(key);
    if (obj === undefined) {
      const paint: ColorPaint = glyph?.paint ?? { kind: 'glyph', gid, paint: { kind: 'solid', color: { rgb: null, alpha: 1 } } };
      obj = this.write(font, paint, glyph?.clip ?? null, fg, opts.group === true);
      byKey.set(key, obj);
    }
    return obj;
  }

  private write(font: SfntFont, paint: ColorPaint, clip: Box | null, fg: string, group: boolean): number | null {
    const [x0, y0, x1, y1] = font.bbox;
    const pad = font.unitsPerEm / 4;
    // The font's box, grown a little: some glyphs reach past it.
    const box: Box = clip ?? [Math.min(x0, 0) - pad, Math.min(y0, 0) - pad, Math.max(x1, font.unitsPerEm) + pad, Math.max(y1, font.unitsPerEm) + pad];
    const res = new Resources();
    const content = new PaintWriter(this.pdf, this.compress, font, fg, box).ops(paint, res, [1, 0, 0, 1, 0, 0]);
    if (!content) return null;
    return this.pdf.addStream(
      `/Type /XObject /Subtype /Form /BBox [${box.map(num).join(' ')}]${group ? ' /Group << /S /Transparency >>' : ''} /Resources ${res}`,
      latin1(content),
      this.compress,
    );
  }
}

class PaintWriter {
  constructor(
    private readonly pdf: PdfObjects,
    private readonly compress: boolean,
    private readonly font: SfntFont,
    private readonly fg: string,
    private readonly box: Box,
  ) {}

  /** Content drawing `p`, where `m` maps the current space to the glyph's. */
  ops(p: ColorPaint, res: Resources, m: Matrix): string {
    switch (p.kind) {
      case 'layers':
        return p.layers.map((l) => this.ops(l, res, m)).join('');
      case 'solid':
        return this.fill(p.color, res, `${this.boxIn(m).map(num).join(' ')} re f `, true);
      case 'glyph': {
        const path = glyphOutline(this.font, p.gid);
        if (!path) return '';
        if (p.paint.kind === 'solid') return this.fill(p.paint.color, res, `${path}f `, false);
        const inner = this.ops(p.paint, res, m);
        return inner ? `q ${path}W n ${inner}Q ` : '';
      }
      case 'linear':
      case 'radial':
        return this.gradient(p, res, m);
      case 'transform': {
        const [a, b, c, d, e, f] = p.matrix;
        if (!Number.isFinite(a + b + c + d + e + f) || Math.abs(a * d - b * c) < 1e-9) return '';
        const inner = this.ops(p.paint, res, multiply(p.matrix, m));
        return inner ? `q ${p.matrix.map(num).join(' ')} cm ${inner}Q ` : '';
      }
      case 'composite':
        return this.composite(p.mode, p.source, p.backdrop, res, m);
    }
  }

  /** The glyph's box in the current space, as x y width height. */
  private boxIn(m: Matrix): Box {
    const inv = invert(m);
    const [l, b, r, t] = this.box;
    const xs: number[] = [];
    const ys: number[] = [];
    for (const [x, y] of [[l, b], [r, b], [l, t], [r, t]]) {
      xs.push(inv[0] * x + inv[2] * y + inv[4]);
      ys.push(inv[1] * x + inv[3] * y + inv[5]);
    }
    const [minX, minY] = [Math.min(...xs), Math.min(...ys)];
    return [minX, minY, Math.max(...xs) - minX, Math.max(...ys) - minY];
  }

  private rgb(c: PaintColor): string {
    const v = parseInt(c.rgb ?? this.fg, 16) || 0;
    return `${num(((v >> 16) & 255) / 255)} ${num(((v >> 8) & 255) / 255)} ${num((v & 255) / 255)}`;
  }

  private fill(c: PaintColor, res: Resources, path: string, save: boolean): string {
    if (c.alpha <= 0) return '';
    if (c.alpha < 1) return `q /${res.state(`<< /ca ${num(c.alpha)} >>`)} gs ${this.rgb(c)} rg ${path}Q `;
    return save ? `q ${this.rgb(c)} rg ${path}Q ` : `${this.rgb(c)} rg ${path}`;
  }

  private gradient(p: Extract<ColorPaint, { kind: 'linear' | 'radial' }>, res: Resources, m: Matrix): string {
    const stops = p.stops;
    if (!stops.length) return '';
    const [o0, o1] = [stops[0].offset, stops[stops.length - 1].offset];
    if (o1 - o0 < 1e-6) return this.ops({ kind: 'solid', color: stops[stops.length - 1].color }, res, m);
    // PDF functions run from 0 to 1, so the geometry moves to where the first and last stops are.
    const at = (a: number, b: number, t: number) => a + (b - a) * t;
    const coords =
      p.kind === 'linear'
        ? [at(p.x0, p.x1, o0), at(p.y0, p.y1, o0), at(p.x0, p.x1, o1), at(p.y0, p.y1, o1)]
        : [at(p.x0, p.x1, o0), at(p.y0, p.y1, o0), Math.max(0, at(p.r0, p.r1, o0)), at(p.x0, p.x1, o1), at(p.y0, p.y1, o1), Math.max(0, at(p.r0, p.r1, o1))];
    const norm = stops.map((s) => ({ offset: (s.offset - o0) / (o1 - o0), color: s.color }));
    const shading = (space: string, value: (s: ColorStop) => string) =>
      `<< /ShadingType ${p.kind === 'linear' ? 2 : 3} /ColorSpace /${space} /Function ${stitch(norm, value)} /Coords [${coords.map(num).join(' ')}] /Extend [true true] >>`;
    const alphas = new Set(norm.map((s) => s.color.alpha));
    const sh = `/${res.shading(shading('DeviceRGB', (s) => this.rgb(s.color)))} sh`;
    if (alphas.size === 1) {
      const a = norm[0].color.alpha;
      if (a <= 0) return '';
      return a < 1 ? `q /${res.state(`<< /ca ${num(a)} >>`)} gs ${sh} Q ` : `q ${sh} Q `;
    }
    // Alpha that varies along the gradient: a soft mask of the same shading in grey.
    const maskRes = new Resources();
    const mask = this.group(`/${maskRes.shading(shading('DeviceGray', (s) => num(s.color.alpha)))} sh`, maskRes, m, '/CS /DeviceGray');
    return `q /${res.state(`<< /SMask << /Type /Mask /S /Luminosity /G ${mask} 0 R >> >>`)} gs ${sh} Q `;
  }

  private composite(mode: number, source: ColorPaint, backdrop: ColorPaint, res: Resources, m: Matrix): string {
    const src = () => this.ops(source, res, m);
    const dst = () => this.ops(backdrop, res, m);
    switch (mode) {
      case COMPOSITE.CLEAR:
        return '';
      case COMPOSITE.SRC:
        return src();
      case COMPOSITE.DEST:
        return dst();
      case COMPOSITE.DEST_OVER:
        return src() + dst();
      case COMPOSITE.SRC_IN:
        return this.masked(source, backdrop, false, res, m);
      case COMPOSITE.DEST_IN:
        return this.masked(backdrop, source, false, res, m);
      case COMPOSITE.SRC_OUT:
        return this.masked(source, backdrop, true, res, m);
      case COMPOSITE.DEST_OUT:
        return this.masked(backdrop, source, true, res, m);
      case COMPOSITE.SRC_ATOP:
        return dst() + this.masked(source, backdrop, false, res, m);
      case COMPOSITE.DEST_ATOP:
        return src() + this.masked(backdrop, source, false, res, m);
      case COMPOSITE.XOR:
        return this.masked(source, backdrop, true, res, m) + this.masked(backdrop, source, true, res, m);
      default: {
        const blend = BLEND_MODES[mode - 13];
        if (!blend) return dst() + src();
        // The source blends with the backdrop only: both sit in a group of their own.
        const srcRes = new Resources();
        const srcForm = this.group(this.ops(source, srcRes, m), srcRes, m, '/I true');
        const groupRes = new Resources();
        const inner = `${this.ops(backdrop, groupRes, m)}q /${groupRes.state(`<< /BM /${blend} >>`)} gs /${groupRes.form(srcForm)} Do Q `;
        return `/${res.form(this.group(inner, groupRes, m, '/I true'))} Do `;
      }
    }
  }

  /** `content` drawn only where `mask` is (or, inverted, isn't), by the mask's alpha. */
  private masked(content: ColorPaint, mask: ColorPaint, invert: boolean, res: Resources, m: Matrix): string {
    const drawn = this.ops(content, res, m);
    if (!drawn) return '';
    const maskRes = new Resources();
    const maskOps = this.ops(mask, maskRes, m);
    if (!maskOps) return invert ? drawn : '';
    const form = this.group(maskOps, maskRes, m, '');
    return `q /${res.state(`<< /SMask << /Type /Mask /S /Alpha /G ${form} 0 R${invert ? ` /TR ${INVERT}` : ''} >> >>`)} gs ${drawn}Q `;
  }

  /** A transparency group form of `content`, in the current space. */
  private group(content: string, res: Resources, m: Matrix, attrs: string): number {
    const [x, y, w, h] = this.boxIn(m);
    return this.pdf.addStream(
      `/Type /XObject /Subtype /Form /BBox [${num(x)} ${num(y)} ${num(x + w)} ${num(y + h)}] /Group << /S /Transparency${attrs ? ` ${attrs}` : ''} >> /Resources ${res}`,
      latin1(content),
      this.compress,
    );
  }
}

/** A PDF function from 0..1 through the stops: linear between each pair, stitched together. */
function stitch(stops: ColorStop[], value: (s: ColorStop) => string): string {
  const parts: string[] = [];
  const bounds: number[] = [];
  for (let i = 0; i + 1 < stops.length; i++) {
    const [a, b] = [stops[i], stops[i + 1]];
    // A stop at the same offset as the next is a hard edge: no segment between them.
    if (b.offset - a.offset < 1e-6) continue;
    if (parts.length) bounds.push(a.offset);
    parts.push(`<< /FunctionType 2 /Domain [0 1] /C0 [${value(a)}] /C1 [${value(b)}] /N 1 >>`);
  }
  if (parts.length === 1) return parts[0];
  return `<< /FunctionType 3 /Domain [0 1] /Functions [${parts.join(' ')}] /Bounds [${bounds.map(num).join(' ')}] /Encode [${parts.map(() => '0 1').join(' ')}] >>`;
}

function invert(m: Matrix): Matrix {
  const det = m[0] * m[3] - m[1] * m[2] || 1;
  const a = m[3] / det;
  const b = -m[1] / det;
  const c = -m[2] / det;
  const d = m[0] / det;
  return [a, b, c, d, -(m[4] * a + m[5] * c), -(m[4] * b + m[5] * d)];
}
