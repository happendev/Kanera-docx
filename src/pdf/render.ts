import type { DocumentInfo, ImageRef, Link } from '../ooxml/model.js';
import type { Compression } from '../util/compression.js';
import type { FontFace } from '../fonts/registry.js';
import { colorGlyph, hasColorGlyphs } from '../fonts/colr.js';
import { buildSfnt, canSubsetTrueType, subsetTrueType } from '../fonts/subset.js';
import { ImageBudget, type ImageLimits } from '../image/budget.js';
import { adjustImage, dibToPdf, sniffImage, toPdfImage, type PdfImage } from '../image/image.js';
import type { DrawOp, GlyphOp, Page, PathSeg, ShadeGeometry } from '../layout/types.js';
import { ColorGlyphForms } from './color-glyph.js';
import { latin1, num, PdfObjects, pdfName, textString } from './objects.js';
import { SYNTHETIC_BOLD_STROKE } from '../profile.js';

export interface PdfOptions {
  info?: DocumentInfo;
  imageLimits?: ImageLimits;
  imageQuality?: number | false;
  additionalLinkSchemes?: readonly string[];
  producer?: string;
  /** Compress content streams (on by default). */
  compress?: boolean;
  /** Snap text sizes, baselines (from the page top) and line starts to this grid, in points. */
  paintGrid?: number;
  /** Leave out, with a warning, a picture past the image pixel limits rather than failing the conversion. */
  skipImagesOverLimits?: boolean;
}

interface FontUse {
  face: FontFace;
  resName: string;
  objNum: number;
  // PDF character codes identify text, while CIDs select outlines. Several characters
  // can share one outline (especially glyph zero when the document font is missing).
  characters: Map<number, string>;
  codesByGlyph: Map<number, Map<string, number>>;
  nextCode: number;
}

interface ImageUse {
  resName: string;
  objNum: number;
}

const ITALIC_SKEW = 0.2126; // tan(12°), the slant Word uses for synthetic italics

/** A number for a matrix that scales font units, where three decimals would lose too much. */
function fine(v: number): string {
  if (!Number.isFinite(v)) return '0';
  const r = Math.round(v * 1e7) / 1e7;
  return r === 0 ? '0' : String(r);
}

/** Write laid-out pages as a PDF file. */
export async function renderPdf(pages: Page[], compression: Compression, opts: PdfOptions = {}, warnings: string[] = []): Promise<Uint8Array<ArrayBuffer>> {
  const quality = opts.imageQuality;
  if (quality !== undefined && quality !== false && (!Number.isInteger(quality) || quality < 1 || quality > 100)) {
    throw new RangeError('Image quality must be an integer from 1 to 100 or false');
  }
  const pdf = new PdfObjects(compression.deflateZlib);
  const catalog = pdf.alloc();
  const pagesNum = pdf.alloc();
  const fonts = new Map<FontFace, FontUse>();
  const imageBudget = new ImageBudget(opts.imageLimits);
  const schemes = new Set(['https', 'http', 'mailto', ...(opts.additionalLinkSchemes ?? []).map((s) => s.toLowerCase())]);
  const images = new Map<string, ImageUse | null>();
  const rasterImages = new WeakMap<Uint8Array, PdfImage | null>();
  const masks = new WeakMap<Uint8Array, number>();
  // Decks often store one picture under several part names: encode and embed it once.
  const byContent = new Map<string, Uint8Array[]>();
  const sameImage = (data: Uint8Array): Uint8Array => {
    const key = contentKey(data);
    const seen = byContent.get(key);
    if (!seen) {
      byContent.set(key, [data]);
      return data;
    }
    for (const other of seen) if (other === data || equalBytes(other, data)) return other;
    seen.push(data);
    return data;
  };
  const writtenImages = new Map<PdfImage, number>();
  // Gradients' sampled colours, shared by gradients that sample the same.
  const shadeFunctions = new Map<string, number>();
  const anchors = new Map<string, { page: number; x: number; y: number }>();
  const pageNums = pages.map(() => pdf.alloc());
  const compress = opts.compress !== false;
  const grid = opts.paintGrid;
  const snap = (v: number) => (grid ? Math.round(v / grid) * grid : v);
  const colorForms = new ColorGlyphForms(pdf, compress);

  // Bookmark destinations first, so links can point forward.
  pages.forEach((page, pi) => {
    for (const layer of [page.behind, page.headerFooter, page.body, page.front])
      for (const op of layer) if (op.kind === 'anchor' && !anchors.has(op.name)) anchors.set(op.name, { page: pi, x: op.x, y: page.height - op.y });
  });

  const fontFor = (face: FontFace): FontUse => {
    let use = fonts.get(face);
    if (!use) {
      use = { face, resName: `F${fonts.size + 1}`, objNum: pdf.alloc(), characters: new Map(), codesByGlyph: new Map(), nextCode: 0 };
      fonts.set(face, use);
    }
    return use;
  };

  const imageFor = (ref: ImageRef): ImageUse | null | undefined => {
    if (ref.svg && images.get(ref.svg.id) !== null) return images.get(ref.svg.id);
    return images.get(ref.id);
  };

  const loadSvg = async (ref: ImageRef): Promise<void> => {
    const { svgToPdfForm } = await import('./svg.js');
    const form = svgToPdfForm(ref.data);
    images.set(ref.id, form === null ? null : {
      resName: `Im${images.size + 1}`,
      objNum: pdf.addStream('/Type /XObject /Subtype /Form /BBox [0 0 1 1] /Resources << >>', latin1(form), compress),
    });
  };

  const loadImage = async (ref: ImageRef): Promise<void> => {
    if (ref.svg) {
      if (!images.has(ref.svg.id)) await loadSvg(ref.svg);
      if (imageFor(ref) !== undefined) return;
    }
    // Raw DIB bits are only an image together with their header, so they are not shared by content.
    const data = ref.dibBits ? ref.data : sameImage(ref.data);
    if (!ref.dibBits && (/\.svg$/i.test(ref.id) || sniffImage(data) === 'svg')) {
      await loadSvg(ref);
      if (!images.get(ref.id)) warnings.push(`Image ${ref.id} uses unsupported SVG features and was skipped`);
      return;
    }
    let img = rasterImages.get(data);
    if (img === undefined) {
      try {
        img = ref.dibBits ? dibToPdf(data, ref.dibBits, compression, imageBudget) : await toPdfImage(data, compression, imageBudget, quality);
      } catch (e) {
        // Past the pixel budget, drop this picture rather than the whole document where asked to.
        if (!opts.skipImagesOverLimits || !(e instanceof Error && e.message === 'Image pixel limit exceeded')) throw e;
        warnings.push(`Image ${ref.id} exceeds the image pixel limits and was skipped`);
        rasterImages.set(data, null);
        images.set(ref.id, null);
        return;
      }
      rasterImages.set(data, img);
    }
    if (!img) {
      warnings.push(`Image ${ref.id} uses an unsupported format and was skipped`);
      images.set(ref.id, null);
      return;
    }
    let objNum = ref.adjust ? undefined : writtenImages.get(img);
    if (objNum === undefined) {
      objNum = writeImage(pdf, ref.adjust ? adjustImage(img, ref.adjust) : img, masks);
      if (!ref.adjust) writtenImages.set(img, objNum);
    }
    images.set(ref.id, { resName: `Im${images.size + 1}`, objNum });
  };

  for (const [pi, page] of pages.entries()) {
    const alphas = new Map<string, string>();
    const alphaFor = (a: number): string => {
      const key = num(a);
      let name = alphas.get(key);
      if (!name) {
        name = `GS${alphas.size + 1}`;
        alphas.set(key, name);
      }
      return name;
    };
    const H = page.height;
    const out: string[] = [];
    const usedFonts = new Set<FontUse>();
    const usedImages = new Set<ImageUse>();
    // Colour glyph forms on this page: object number → resource name.
    const usedForms = new Map<number, string>();
    const formName = (obj: number): string => {
      let name = usedForms.get(obj);
      if (!name) usedForms.set(obj, (name = `Cg${usedForms.size + 1}`));
      return name;
    };
    const shadings: string[] = [];
    const softMasks: string[] = [];
    const annots: string[] = [];
    const Y = (y: number) => H - y;
    const color = (hex: string, stroke: boolean) => {
      const v = parseInt(hex.slice(0, 6), 16) || 0;
      return `${num(((v >> 16) & 255) / 255)} ${num(((v >> 8) & 255) / 255)} ${num((v & 255) / 255)} ${stroke ? 'RG' : 'rg'}`;
    };
    const pathOps = (segs: PathSeg[]): string => {
      let s = '';
      for (const seg of segs) {
        const p = seg.pts;
        if (seg.op === 'M') s += `${num(p[0])} ${num(Y(p[1]))} m `;
        else if (seg.op === 'L') s += `${num(p[0])} ${num(Y(p[1]))} l `;
        else if (seg.op === 'C') s += `${num(p[0])} ${num(Y(p[1]))} ${num(p[2])} ${num(Y(p[3]))} ${num(p[4])} ${num(Y(p[5]))} c `;
        else s += 'h ';
      }
      return s;
    };

    const glyphs = (op: GlyphOp) => {
      if (!op.gids.length) return;
      const use = fontFor(op.face);
      usedFonts.add(use);
      const font = op.face.font;
      const toPdfUnits = 1000 / font.unitsPerEm;
      const th = op.scaleX || 1;
      // A colour font's glyphs draw as forms over invisible text, which keeps the text searchable.
      const colored = hasColorGlyphs(font) && op.gids.some((g) => colorGlyph(font, g) !== null);
      // Text render mode, horizontal scaling, stroke colour and line width outlive ET, so stroked
      // (bold or outline), invisible and scaled text saves and restores the graphics state;
      // otherwise every later run would be stroked, invisible or scaled too.
      const stroked = !colored && (op.syntheticBold || op.outline);
      const alpha = op.alpha !== undefined && op.alpha < 1;
      const save = alpha || stroked || colored || th !== 1;
      if (alpha) out.push(`q /${alphaFor(op.alpha!)} gs`);
      else if (save) out.push('q');
      // On a paint grid only the drawn size snaps; glyphs keep their laid-out positions.
      const size = grid ? Math.max(grid, snap(op.size)) : op.size;
      let s = `BT /${use.resName} ${num(size)} Tf `;
      if (th !== 1) s += `${num(th * 100)} Tz `;
      s += color(op.color, false) + ' ';
      if (stroked) {
        // A dash pattern a line or border left set would dash the glyphs' outlines too.
        s += `${color(op.color, true)} ${num(op.size * (op.outline ? 0.02 : SYNTHETIC_BOLD_STROKE))} w [] 0 d ${op.outline ? 1 : 2} Tr `;
      }
      if (colored) s += '3 Tr ';
      const x0 = op.x + op.xs[0] + (grid && op.lineX !== undefined ? snap(op.lineX) - op.lineX : 0);
      s += `1 0 ${op.syntheticItalic ? ITALIC_SKEW : 0} 1 ${num(x0)} ${num(Y(grid ? snap(op.y) : op.y))} Tm [`;
      for (let i = 0; i < op.gids.length; i++) {
        const gid = op.gids[i];
        const code = characterCode(use, gid, op.texts[i] ?? '');
        s += `<${code.toString(16).padStart(4, '0')}>`;
        if (size !== 0 && i + 1 < op.gids.length) {
          const w = font.advance(gid) * toPdfUnits;
          const d = op.xs[i + 1] - op.xs[i];
          const adj = w - (d * 1000) / (size * th);
          if (Math.abs(adj) > 0.05) s += num(adj);
        }
      }
      s += '] TJ ET';
      out.push(s);
      if (colored) {
        const k = size / font.unitsPerEm;
        const skew = op.syntheticItalic ? ITALIC_SKEW : 0;
        const y = Y(grid ? snap(op.y) : op.y);
        for (let i = 0; i < op.gids.length; i++) {
          const form = colorForms.form(font, op.gids[i], op.color, { plain: true, group: alpha });
          if (form === null) continue;
          out.push(`q ${fine(k * th)} 0 ${fine(k * skew)} ${fine(k)} ${num(x0 + op.xs[i] - op.xs[0])} ${num(y)} cm /${formName(form)} Do Q`);
        }
      }
      if (save) out.push('Q');
    };

    const draw = (op: DrawOp) => {
      switch (op.kind) {
        case 'glyphs':
          glyphs(op);
          break;
        case 'rect': {
          const a = op.alpha !== undefined && op.alpha < 1;
          out.push(`${a ? `q /${alphaFor(op.alpha!)} gs ` : ''}${color(op.fill, false)} ${num(op.x)} ${num(Y(op.y + op.h))} ${num(op.w)} ${num(op.h)} re f${a ? ' Q' : ''}`);
          break;
        }
        case 'line': {
          const cap = op.cap === 'round' ? 1 : op.cap === 'square' ? 2 : 0;
          const dash = op.dash ? `[${op.dash.map(num).join(' ')}] 0 d` : '[] 0 d';
          out.push(`${num(op.width)} w ${cap} J ${dash} ${color(op.color, true)} ${num(op.x1)} ${num(Y(op.y1))} m ${num(op.x2)} ${num(Y(op.y2))} l S`);
          break;
        }
        case 'path': {
          const path = pathOps(op.segs);
          if (op.fill) {
            const a = op.fillAlpha !== undefined && op.fillAlpha < 1;
            out.push(`${a ? `q /${alphaFor(op.fillAlpha!)} gs ` : ''}${color(op.fill, false)} ${path}${op.nonzero ? 'f' : 'f*'}${a ? ' Q' : ''}`);
          }
          if (op.stroke) {
            const a = op.strokeAlpha !== undefined && op.strokeAlpha < 1;
            const dash = op.dash ? `[${op.dash.map(num).join(' ')}] 0 d` : '[] 0 d';
            const cap = op.cap === 'round' ? 1 : op.cap === 'square' ? 2 : 0;
            const join = op.join === 'miter' ? 0 : op.join === 'bevel' ? 2 : 1;
            out.push(`${a ? `q /${alphaFor(op.strokeAlpha!)} gs ` : ''}${num(op.strokeWidth ?? 0.75)} w ${cap} J ${join} j ${dash} ${color(op.stroke, true)} ${path}S${a ? ' Q' : ''}`);
          }
          break;
        }
        case 'shade': {
          const g = op.geometry;
          const fn = sampledFunction(pdf, shadeFunctions, op.colors.flatMap(rgbBytes), 3, compress);
          const name = `Sh${shadings.length + 1}`;
          shadings.push(`/${name} ${shadingObject(pdf, g, fn, 'DeviceRGB', Y, compress)} 0 R`);
          // A radial shading is a circle's, scaled to the ellipse.
          const cm = g.type === 'radial' ? `${num(g.rx)} 0 0 ${num(g.ry)} ${num(g.cx)} ${num(Y(g.cy))} cm ` : '';
          let s = 'q ';
          if (op.alphas) {
            // Opacity that varies across the gradient is a soft mask: the same shading in grey.
            const maskFn = sampledFunction(pdf, shadeFunctions, op.alphas.map((a) => Math.round(Math.max(0, Math.min(1, a)) * 255)), 1, compress);
            const [l, t, r, b] = op.bbox;
            const form = pdf.addStream(
              `/Type /XObject /Subtype /Form /BBox [${num(l - 1)} ${num(Y(b) - 1)} ${num(r + 1)} ${num(Y(t) + 1)}] /Group << /S /Transparency /CS /DeviceGray >> /Resources << /Shading << /Sh ${shadingObject(pdf, g, maskFn, 'DeviceGray', Y, compress)} 0 R >> >>`,
              latin1(`${cm}/Sh sh`),
              compress,
            );
            const mask = `SM${softMasks.length + 1}`;
            softMasks.push(`/${mask} ${pdf.add(`<< /Type /ExtGState /SMask << /Type /Mask /S /Luminosity /G ${form} 0 R >> >>`)} 0 R`);
            s += `/${mask} gs `;
          } else if (op.alpha !== undefined && op.alpha < 1) s += `/${alphaFor(op.alpha)} gs `;
          out.push(`${s}${cm}/${name} sh Q`);
          break;
        }
        case 'image': {
          const use = imageFor(op.image);
          if (!use) break;
          usedImages.add(use);
          const c = op.crop ?? { l: 0, t: 0, r: 0, b: 0 };
          const fw = 1 - c.l - c.r;
          const fh = 1 - c.t - c.b;
          if (fw <= 0 || fh <= 0) break;
          const W = op.w / fw;
          const Hh = op.h / fh;
          const ix = op.x - c.l * W;
          const iy = op.y - c.t * Hh;
          let s = 'q ';
          if (op.alpha !== undefined && op.alpha < 1) s += `/${alphaFor(op.alpha)} gs `;
          if (c.l || c.t || c.r || c.b) s += `${num(op.x)} ${num(Y(op.y + op.h))} ${num(op.w)} ${num(op.h)} re W n `;
          s += `${num(W)} 0 0 ${num(Hh)} ${num(ix)} ${num(Y(iy + Hh))} cm /${use.resName} Do Q`;
          out.push(s);
          break;
        }
        case 'transform': {
          const [a, b, c, d, e, f] = op.matrix;
          out.push(`q ${num(a)} ${num(-b)} ${num(-c)} ${num(d)} ${num(c * H + e)} ${num(H - d * H - f)} cm`);
          break;
        }
        case 'clip':
          if (op.segs) out.push(`q ${pathOps(op.segs)}${op.nonzero ? 'W' : 'W*'} n`);
          else out.push(`q ${num(op.x)} ${num(Y(op.y + op.h))} ${num(op.w)} ${num(op.h)} re W n`);
          break;
        case 'restore':
          out.push('Q');
          break;
        case 'link': {
          const action = linkAction(op.link, anchors, pageNums, schemes);
          if (!action) break;
          annots.push(`<< /Type /Annot /Subtype /Link /Rect [${num(op.x)} ${num(Y(op.y + op.h))} ${num(op.x + op.w)} ${num(Y(op.y))}] /Border [0 0 0] ${action} >>`);
          break;
        }
        case 'anchor':
        case 'float':
          break;
      }
    };
    for (const layer of [page.behind, page.headerFooter, page.body, page.front]) {
      for (const op of layer) {
        if (op.kind === 'image' && imageFor(op.image) === undefined) await loadImage(op.image);
        draw(op);
      }
    }

    const content = pdf.addStream('', latin1(out.join('\n')), compress);
    const fontRes = [...usedFonts].map((f) => `/${f.resName} ${f.objNum} 0 R`).join(' ');
    const imgRes = [...[...usedImages].map((i) => `/${i.resName} ${i.objNum} 0 R`), ...[...usedForms].map(([obj, name]) => `/${name} ${obj} 0 R`)].join(' ');
    const gsRes = [...[...alphas].map(([a, n]) => `/${n} << /Type /ExtGState /ca ${a} /CA ${a} >>`), ...softMasks].join(' ');
    const shRes = shadings.join(' ');
    const annotNums = annots.map((a) => pdf.add(a));
    pdf.set(
      pageNums[pi],
      `<< /Type /Page /Parent ${pagesNum} 0 R /MediaBox [0 0 ${num(page.width)} ${num(page.height)}] /Resources << /ProcSet [/PDF /Text /ImageB /ImageC /ImageI]${fontRes ? ` /Font << ${fontRes} >>` : ''}${imgRes ? ` /XObject << ${imgRes} >>` : ''}${gsRes ? ` /ExtGState << ${gsRes} >>` : ''}${shRes ? ` /Shading << ${shRes} >>` : ''} >> /Contents ${content} 0 R${annotNums.length ? ` /Annots [${annotNums.map((n) => `${n} 0 R`).join(' ')}]` : ''} >>`,
    );
  }

  let fi = 0;
  for (const use of fonts.values()) {
    const gids = [...use.codesByGlyph.keys()].sort((a, b) => a - b);
    const read = use.face.font.readSubsetGlyphs;
    writeFont(pdf, use, fi++, compress, gids, read ? await read(gids) : undefined);
  }

  pdf.set(pagesNum, `<< /Type /Pages /Kids [${pageNums.map((n) => `${n} 0 R`).join(' ')}] /Count ${pages.length} >>`);
  pdf.set(catalog, `<< /Type /Catalog /Pages ${pagesNum} 0 R >>`);
  const info = opts.info ?? {};
  const entries = [`/Producer ${textString(opts.producer ?? '@kanera/docx2pdf')}`];
  if (info.title) entries.push(`/Title ${textString(info.title)}`);
  if (info.creator) entries.push(`/Author ${textString(info.creator)}`);
  if (info.subject) entries.push(`/Subject ${textString(info.subject)}`);
  if (info.keywords) entries.push(`/Keywords ${textString(info.keywords)}`);
  // PDF dates are D:YYYYMMDDHHmmSS, here in UTC (Z); a year outside 0000-9999 has no form.
  const pdfDate = (d: Date | undefined) => {
    const iso = d instanceof Date && !isNaN(d.getTime()) ? d.toISOString() : '';
    return /^\d{4}-/.test(iso) ? `(D:${iso.replace(/\D/g, '').slice(0, 14)}Z)` : undefined;
  };
  const created = pdfDate(info.created);
  const modified = pdfDate(info.modified);
  if (created) entries.push(`/CreationDate ${created}`);
  if (modified) entries.push(`/ModDate ${modified}`);
  const infoNum = pdf.add(`<< ${entries.join(' ')} >>`);
  return pdf.finish(catalog, infoNum);
}

function linkAction(link: Link, anchors: Map<string, { page: number; x: number; y: number }>, pageNums: number[], schemes: ReadonlySet<string>): string | null {
  const bookmark = link.url ? (link.url.startsWith('#') ? link.url.slice(1) : undefined) : link.anchor;
  if (bookmark) {
    const a = anchors.get(bookmark);
    if (a) return `/Dest [${pageNums[a.page]} 0 R /XYZ ${num(a.x)} ${num(a.y)} 0]`;
  }
  if (link.url) {
    let url = link.url;
    if (link.anchor) url += '#' + link.anchor;
    const scheme = /^([a-z][a-z0-9+.-]*):/i.exec(url)?.[1].toLowerCase();
    if (!scheme || !schemes.has(scheme) || /[\x00-\x1f\x7f\uD800-\uDFFF]/u.test(url)) return null;
    const ascii = url.replace(/[^\x20-\x7e]/gu, (c) => encodeURIComponent(c)).replace(/[\\()]/g, (c) => '\\' + c);
    return `/A << /S /URI /URI (${ascii}) >>`;
  }
  return null;
}

const rgbBytes = (hex: string): number[] => {
  const v = parseInt(hex.slice(0, 6), 16) || 0;
  return [(v >> 16) & 255, (v >> 8) & 255, v & 255];
};

/** A sampled function (FunctionType 0) from 0..1 to `samples`, `comps` bytes per evenly spaced step. */
function sampledFunction(pdf: PdfObjects, cache: Map<string, number>, samples: number[], comps: number, compress: boolean): number {
  const key = `${comps}:${samples.join(',')}`;
  let n = cache.get(key);
  if (n === undefined) {
    const range = Array.from({ length: comps }, () => '0 1').join(' ');
    n = pdf.addStream(`/FunctionType 0 /Domain [0 1] /Range [${range}] /Size [${samples.length / comps}] /BitsPerSample 8`, Uint8Array.from(samples), compress);
    cache.set(key, n);
  }
  return n;
}

/**
 * A shading of `geometry` (page coordinates, flipped by `Y`) coloured by `fn`. A radial one is the
 * unit circle's, to be scaled to its ellipse; a rectangular one is a mesh of triangles from the
 * inner rectangle's edges (at 0) to the outer's (at 1), along which the rings' colour runs straight.
 */
function shadingObject(pdf: PdfObjects, g: ShadeGeometry, fn: number, space: string, Y: (y: number) => number, compress: boolean): number {
  const head = `/ColorSpace /${space} /Function ${fn} 0 R`;
  if (g.type === 'axial') return pdf.add(`<< /ShadingType 2 ${head} /Coords [${num(g.x0)} ${num(Y(g.y0))} ${num(g.x1)} ${num(Y(g.y1))}] /Extend [true true] >>`);
  if (g.type === 'radial') return pdf.add(`<< /ShadingType 3 ${head} /Coords [0 0 0 0 0 1] /Extend [true true] >>`);
  const [ol, ot, or, ob] = g.outer;
  const [il, it, ir, ib] = g.inner;
  const [x0, x1, y0, y1] = [ol, or, Y(ob), Y(ot)];
  const O = [[ol, ot], [or, ot], [or, ob], [ol, ob]];
  const I = [[il, it], [ir, it], [ir, ib], [il, ib]];
  const triangles: Array<Array<[number[], number]>> = [];
  for (let k = 0; k < 4; k++) {
    const [o0, o1, i0, i1] = [O[k], O[(k + 1) % 4], I[k], I[(k + 1) % 4]];
    triangles.push([[o0, 1], [o1, 1], [i1, 0]], [[o0, 1], [i1, 0], [i0, 0]]);
  }
  triangles.push([[I[0], 0], [I[1], 0], [I[2], 0]], [[I[0], 0], [I[2], 0], [I[3], 0]]);
  // Each vertex: an edge flag byte, x and y as 32-bit fractions of the box, its position as a 16-bit one.
  const data = new Uint8Array(triangles.length * 3 * 11);
  const view = new DataView(data.buffer);
  const frac = (v: number, lo: number, hi: number) => Math.round(Math.max(0, Math.min(1, (v - lo) / (hi - lo || 1))) * 0xffffffff);
  let p = 0;
  for (const tri of triangles) {
    for (const [[x, y], t] of tri) {
      view.setUint8(p, 0);
      view.setUint32(p + 1, frac(x, x0, x1));
      view.setUint32(p + 5, frac(Y(y), y0, y1));
      view.setUint16(p + 9, t * 0xffff);
      p += 11;
    }
  }
  return pdf.addStream(`/ShadingType 4 ${head} /BitsPerCoordinate 32 /BitsPerComponent 16 /BitsPerFlag 8 /Decode [${num(x0)} ${num(x1)} ${num(y0)} ${num(y1)} 0 1]`, data, compress);
}

/** Length and a sparse hash: equal keys are confirmed byte for byte. */
function contentKey(d: Uint8Array): string {
  let h = 0x811c9dc5;
  const step = Math.max(1, Math.floor(d.length / 512));
  for (let i = 0; i < d.length; i += step) h = Math.imul(h ^ d[i], 0x01000193);
  for (let i = Math.max(0, d.length - 64); i < d.length; i++) h = Math.imul(h ^ d[i], 0x01000193);
  return `${d.length}:${h >>> 0}`;
}

function equalBytes(a: Uint8Array, b: Uint8Array): boolean {
  if (a.length !== b.length) return false;
  for (let i = 0; i < a.length; i++) if (a[i] !== b[i]) return false;
  return true;
}

function writeImage(pdf: PdfObjects, img: PdfImage, masks: WeakMap<Uint8Array, number>): number {
  let smask = '';
  if (img.smask) {
    let n = masks.get(img.smask);
    if (n === undefined) {
      n = pdf.addStream(`/Type /XObject /Subtype /Image /Width ${img.width} /Height ${img.height} /ColorSpace /DeviceGray /BitsPerComponent 8 /Filter /FlateDecode`, img.smask, false);
      masks.set(img.smask, n);
    }
    smask = ` /SMask ${n} 0 R`;
  }
  const parms = img.decodeParms ? ` /DecodeParms ${img.decodeParms}` : '';
  const decode = img.decode ? ` /Decode [${img.decode.join(' ')}]` : '';
  return pdf.addStream(
    `/Type /XObject /Subtype /Image /Width ${img.width} /Height ${img.height} /ColorSpace ${img.colorSpace} /BitsPerComponent ${img.bitsPerComponent} /Filter /${img.filter}${parms}${decode}${smask}`,
    img.data,
    false,
  );
}

function subsetTag(i: number): string {
  let s = '';
  let n = i + 1;
  for (let k = 0; k < 6; k++) {
    s = String.fromCharCode(65 + (n % 26)) + s;
    n = Math.floor(n / 26);
  }
  return s;
}

function characterCode(use: FontUse, gid: number, text: string): number {
  let codes = use.codesByGlyph.get(gid);
  if (!codes) use.codesByGlyph.set(gid, (codes = new Map()));
  const known = codes.get(text);
  if (known !== undefined) return known;
  let code = gid;
  if (use.characters.has(code)) {
    while (use.characters.has(use.nextCode)) use.nextCode++;
    code = use.nextCode++;
    if (code > 0xffff) throw new Error('Too many distinct characters for a PDF font');
  }
  codes.set(text, code);
  use.characters.set(code, text);
  return code;
}

function fontEncoding(pdf: PdfObjects, use: FontUse, compress: boolean): string {
  const pairs: Array<[number, number]> = [];
  let identity = true;
  for (const [gid, codes] of use.codesByGlyph) {
    for (const code of codes.values()) {
      pairs.push([code, gid]);
      if (code !== gid) identity = false;
    }
  }
  if (identity) return '/Identity-H';
  const info = '/CIDSystemInfo << /Registry (Adobe) /Ordering (Identity) /Supplement 0 >>';
  const name = '/' + use.resName + '-Encoding';
  const lines = ['/CIDInit /ProcSet findresource begin 12 dict begin begincmap',
    info + ' def', '/CMapName ' + name + ' def /CMapType 1 def',
    '1 begincodespacerange <0000> <FFFF> endcodespacerange'];
  for (let i = 0; i < pairs.length; i += 100) {
    const chunk = pairs.slice(i, i + 100);
    lines.push(chunk.length + ' begincidchar');
    for (const [code, gid] of chunk) lines.push('<' + code.toString(16).padStart(4, '0') + '> ' + gid);
    lines.push('endcidchar');
  }
  lines.push('endcmap CMapName currentdict /CMap defineresource pop end end');
  const obj = pdf.addStream('/Type /CMap /CMapName ' + name + ' ' + info, latin1(lines.join('\n')), compress);
  return obj + ' 0 R';
}

function writeFont(pdf: PdfObjects, use: FontUse, index: number, compress: boolean, gids: number[], outlines?: ReadonlyMap<number, Uint8Array>): void {
  const font = use.face.font;
  const upm = font.unitsPerEm;
  const scale = 1000 / upm;
  const psName = (font.names.postscriptName || use.face.familyName.replace(/\s+/g, '')).replace(/[^\x21-\x7e]/g, '') || use.resName;
  const canSubset = canSubsetTrueType(font);
  const baseFont = canSubset ? subsetTag(index) + '+' + psName : psName;

  let fileKey: string;
  let fileNum: number;
  let subtype: string;
  let cidToGidMap = '/Identity';
  if (canSubset) {
    const { data, glyphIds } = subsetTrueType(font, gids, outlines);
    fileNum = pdf.addStream(`/Length1 ${data.length}`, data, compress);
    fileKey = 'FontFile2';
    subtype = 'CIDFontType2';
    // Keep the PDF's character codes and CIDs stable, including aliases for shared outlines.
    if (gids.some((gid) => glyphIds.get(gid) !== gid)) {
      const map = new Uint8Array(2 * (gids[gids.length - 1] + 1));
      const entries = new DataView(map.buffer);
      for (const gid of gids) entries.setUint16(gid * 2, glyphIds.get(gid)!);
      cidToGidMap = pdf.addStream('', map, compress) + ' 0 R';
    }
  } else {
    // CFF and fonts we cannot subset are embedded as standalone font programs.
    let data = font.data;
    if (font.dirOffset !== 0) {
      const tables = new Map<string, Uint8Array>();
      for (const tag of font.tables.keys()) tables.set(tag, font.tableBytes(tag)!);
      data = buildSfnt(tables, font.sfntVersion);
    }
    fileNum = pdf.addStream(font.isCFF ? '/Subtype /OpenType' : '/Length1 ' + data.length, data, compress);
    fileKey = font.isCFF ? 'FontFile3' : 'FontFile2';
    subtype = font.isCFF ? 'CIDFontType0' : 'CIDFontType2';
  }

  // Widths
  let w = '';
  for (let i = 0; i < gids.length; ) {
    let j = i;
    const run: string[] = [];
    while (j < gids.length && gids[j] === gids[i] + (j - i)) {
      run.push(num(font.advance(gids[j]) * scale));
      j++;
    }
    w += `${gids[i]} [${run.join(' ')}] `;
    i = j;
  }

  let flags = 0;
  if (font.isFixedPitch) flags |= 1;
  flags |= 32; // nonsymbolic
  if (use.face.italic) flags |= 64;
  const [x0, y0, x1, y1] = font.bbox;
  const descriptor = pdf.add(
    `<< /Type /FontDescriptor /FontName ${pdfName(baseFont)} /Flags ${flags} /FontBBox [${num(x0 * scale)} ${num(y0 * scale)} ${num(x1 * scale)} ${num(y1 * scale)}] /ItalicAngle ${num(font.italicAngle)} /Ascent ${num(font.hheaAscender * scale)} /Descent ${num(font.hheaDescender * scale)} /CapHeight ${num((font.capHeight || font.hheaAscender * 0.7) * scale)} /StemV ${use.face.bold ? 120 : 80} /${fileKey} ${fileNum} 0 R >>`,
  );
  const cid = pdf.add(
    `<< /Type /Font /Subtype /${subtype} /BaseFont ${pdfName(baseFont)} /CIDSystemInfo << /Registry (Adobe) /Ordering (Identity) /Supplement 0 >> /FontDescriptor ${descriptor} 0 R /DW ${num(font.advance(0) * scale)} /W [${w}]${subtype === 'CIDFontType2' ? ' /CIDToGIDMap ' + cidToGidMap : ''} >>`,
  );
  const toUnicode = pdf.addStream('', latin1(toUnicodeCMap(use.characters)), compress);
  pdf.set(use.objNum, `<< /Type /Font /Subtype /Type0 /BaseFont ${pdfName(baseFont)} /Encoding ${fontEncoding(pdf, use, compress)} /DescendantFonts [${cid} 0 R] /ToUnicode ${toUnicode} 0 R >>`);
}

function utf16Hex(s: string): string {
  let h = '';
  for (let i = 0; i < s.length; i++) h += s.charCodeAt(i).toString(16).padStart(4, '0');
  return h.toUpperCase();
}

function toUnicodeCMap(glyphs: Map<number, string>): string {
  const entries = [...glyphs].filter(([, t]) => t.length > 0).sort((a, b) => a[0] - b[0]);
  let s = '/CIDInit /ProcSet findresource begin\n12 dict begin\nbegincmap\n/CIDSystemInfo << /Registry (Adobe) /Ordering (UCS) /Supplement 0 >> def\n/CMapName /Adobe-Identity-UCS def\n/CMapType 2 def\n1 begincodespacerange\n<0000> <FFFF>\nendcodespacerange\n';
  for (let i = 0; i < entries.length; i += 100) {
    const chunk = entries.slice(i, i + 100);
    s += `${chunk.length} beginbfchar\n`;
    for (const [code, text] of chunk) s += `<${code.toString(16).padStart(4, '0').toUpperCase()}> <${utf16Hex(text)}>\n`;
    s += 'endbfchar\n';
  }
  s += 'endcmap\nCMapName currentdict /CMap defineresource pop\nend\nend\n';
  return s;
}
