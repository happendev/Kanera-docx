import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import { colorGlyph, glyphOutline, hasColorGlyphs, type ColorPaint } from '../src/fonts/colr.js';
import { emojiSequence } from '../src/fonts/emoji.js';
import { FontRegistry } from '../src/fonts/registry.js';
import { SfntFont } from '../src/fonts/sfnt.js';
import { buildSfnt } from '../src/fonts/subset.js';
import { shapeText } from '../src/layout/shaper.js';
import { DEFAULT_RUN } from '../src/docx/props.js';
import { convertDocx } from '../src/node.js';
import { ghostscript, ghostscriptPixels, pdfPagesText, pdfText } from './helpers/pdf.js';
import { makeDocx, p } from './helpers/docx.js';

const cps = (s: string) => Array.from(s, (c) => c.codePointAt(0)!);
const font = (file: string) => readFileSync(join(import.meta.dirname, '../fonts', file));

describe('emoji sequences', () => {
  it('cover a character that is an emoji by default, or asked for as one', () => {
    expect(emojiSequence(cps('😀a'), 0)).toBe(1);
    expect(emojiSequence(cps('☀️'), 0)).toBe(2);
    expect(emojiSequence(cps('👍🏻'), 0)).toBe(2);
    expect(emojiSequence(cps('👨‍👩‍👧'), 0)).toBe(5);
    expect(emojiSequence(cps('1️⃣'), 0)).toBe(3);
    expect(emojiSequence(cps('🇿🇦'), 0)).toBe(2);
  });

  it('leave out symbols that are text by default and text asked for as text', () => {
    expect(emojiSequence(cps('☀'), 0)).toBe(0);
    expect(emojiSequence(cps('❄ '), 0)).toBe(0);
    expect(emojiSequence(cps('😀︎'), 0)).toBe(0);
    expect(emojiSequence(cps('★'), 0)).toBe(0);
    expect(emojiSequence(cps('1'), 0)).toBe(0);
  });
});

describe('emoji shaping', () => {
  const fonts = new FontRegistry();
  fonts.add({ data: font('Carlito-Regular.ttf') }, 'bundled');
  fonts.add({ data: font('DejaVuSans.ttf') }, 'bundled');
  fonts.add({ data: font('Noto-COLRv1.ttf') }, 'bundled');
  const props = { ...DEFAULT_RUN, fontAscii: 'Calibri', fontHAnsi: 'Calibri', fontEastAsia: 'Calibri', fontCs: 'Calibri', size: 11, sizeCs: 11 };
  const shape = (text: string, extra = {}) =>
    shapeText(text, { ...props, ...extra }, fonts).flatMap((run) => run.glyphs.map((g) => ({ family: run.face.face.familyName, size: run.size, shift: run.baselineShift, top: run.ascent + run.baselineShift, bottom: run.descent - run.baselineShift, ...g })));

  it('joins a person and skin tone, a family and a flag into one glyph each', () => {
    const glyphs = shape('👍🏻👨‍👩‍👧🇿🇦').filter((g) => g.gid >= 0);
    expect(glyphs.map((g) => g.text)).toEqual(['👍🏻', '👨‍👩‍👧', '🇿🇦']);
    expect(glyphs.every((g) => g.family === 'Noto Color Emoji')).toBe(true);
  });

  it('draws a symbol from the emoji font only when asked for an emoji', () => {
    const emoji = shape('☀️');
    expect(emoji[0].family).toBe('Noto Color Emoji');
    expect(emoji[0].text).toBe('☀️');
    // Without U+FE0F a symbol keeps the text's font where it has one, else the symbol fallback.
    expect(shape('☀')[0].family).toBe('Carlito');
    expect(shape('❄')[0].family).toBe('DejaVu Sans');
  });

  it('draws Noto Color Emoji at the size and in the space of Segoe UI Emoji\'s, on the text\'s line', () => {
    const [text, plain] = shape('a😀');
    // Glyphs 92% of the text's size, raised 0.03 em, centred in Segoe's 1.373 em advance.
    expect(plain.size).toBeCloseTo(11 * 0.92, 6);
    expect(plain.shift).toBeCloseTo(0.03 * 11, 6);
    expect(plain.advance).toBeCloseTo(1.373 * 11, 6);
    expect(plain.dx).toBeCloseTo((1.373 * 11 - (1275 / 1024) * 11 * 0.92) / 2, 6);
    // Moving the glyphs leaves the line as tall as the text makes it.
    expect([plain.top, plain.bottom]).toEqual([text.top, text.bottom]);
    const [raised] = shape('😀', { position: 3 });
    expect(raised.shift).toBeCloseTo(3 + 0.03 * 11, 6);
  });

  it('draws nothing for a variation selector or joiner the font passes over', () => {
    const glyphs = shape('a\u{FE0F}b‍c');
    expect(glyphs.filter((g) => g.gid >= 0).map((g) => g.text).join('')).toBe('abc');
    expect(glyphs.filter((g) => g.gid < 0).every((g) => g.advance === 0)).toBe(true);
  });
});

describe('colour glyphs', () => {
  const noto = new SfntFont(font('Noto-COLRv1.ttf'));

  it('reads COLR version 1 paint graphs: layers, gradients, transforms and clip boxes', () => {
    const smile = colorGlyph(noto, noto.glyphId(0x1f60a))!;
    const kinds = new Set<string>();
    const walk = (p: ColorPaint): void => {
      kinds.add(p.kind);
      if (p.kind === 'layers') p.layers.forEach(walk);
      else if (p.kind === 'glyph' || p.kind === 'transform') walk(p.paint);
      else if (p.kind === 'composite') [p.source, p.backdrop].forEach(walk);
    };
    walk(smile.paint);
    expect([...kinds].sort()).toEqual(['glyph', 'layers', 'radial', 'solid', 'transform']);
    expect(smile.clip).toEqual([64, -192, 1184, 896]);
    expect(smile.foreground).toBe(false);
    expect(colorGlyph(noto, 0)).toBeNull();
  });

  it('turns TrueType outlines into closed PDF paths, components included', () => {
    const carlito = new SfntFont(font('Carlito-Regular.ttf'));
    const contours = (cp: number) => glyphOutline(carlito, carlito.glyphId(cp)).match(/ h /g)?.length;
    expect(contours(0x4f)).toBe(2); // O
    expect(contours(0xe9)).toBe((contours(0x65) ?? 0) + 1); // é: e and the accent
    expect(glyphOutline(carlito, carlito.glyphId(0x20))).toBe('');
    expect(hasColorGlyphs(carlito)).toBe(false);
  });

  it('draws COLR version 0 layers in palette colours and the text colour, over searchable text', async () => {
    // Carlito with a colour "A": a red "O" layer under an "I" in the text's colour.
    const carlito = new SfntFont(font('Carlito-Regular.ttf'));
    const [A, O, I] = [0x41, 0x4f, 0x49].map((cp) => carlito.glyphId(cp));
    const colr = new DataView(new ArrayBuffer(28));
    colr.setUint16(2, 1); // one base glyph record, at 14
    colr.setUint32(4, 14);
    colr.setUint32(8, 20); // two layer records, at 20
    colr.setUint16(12, 2);
    colr.setUint16(14, A);
    colr.setUint16(18, 2);
    colr.setUint16(20, O);
    colr.setUint16(24, I);
    colr.setUint16(26, 0xffff);
    const cpal = new Uint8Array([0, 0, 0, 1, 0, 1, 0, 1, 0, 0, 0, 14, 0, 0, 0, 0, 0xff, 0xff]);
    const tables = new Map([...carlito.tables.keys()].map((tag) => [tag, carlito.tableBytes(tag)!]));
    tables.set('COLR', new Uint8Array(colr.buffer));
    tables.set('CPAL', cpal);
    const data = buildSfnt(tables, carlito.sfntVersion);
    expect(colorGlyph(new SfntFont(data), A)?.paint).toEqual({
      kind: 'layers',
      layers: [
        { kind: 'glyph', gid: O, paint: { kind: 'solid', color: { rgb: 'FF0000', alpha: 1 } } },
        { kind: 'glyph', gid: I, paint: { kind: 'solid', color: { rgb: null, alpha: 1 } } },
      ],
    });
    const body = '<w:p><w:r><w:rPr><w:rFonts w:ascii="Carlito" w:hAnsi="Carlito"/><w:color w:val="0000FF"/><w:sz w:val="144"/></w:rPr><w:t>AB</w:t></w:r></w:p>';
    const result = await convertDocx(makeDocx({ body }), { compress: false, fonts: [{ data }] });
    expect(pdfText(result.pdf)).toMatch(/3 Tr .*\] TJ ET\nq [\d.]+ 0 0 [\d.]+ [\d.]+ [\d.]+ cm \/Cg1 Do Q/);
    if (!ghostscript) return;
    expect(pdfPagesText(result.pdf)[0]).toContain('AB');
    const { rgb } = ghostscriptPixels(result.pdf, 36);
    let [red, blue] = [0, 0];
    for (let i = 0; i < rgb.length; i += 3) {
      if (rgb[i] > 200 && rgb[i + 1] < 80 && rgb[i + 2] < 80) red++;
      if (rgb[i] < 80 && rgb[i + 1] < 80 && rgb[i + 2] > 200) blue++;
    }
    expect(red).toBeGreaterThan(20);
    expect(blue).toBeGreaterThan(20);
  });
});

describe('converting a document with emoji', () => {
  it('draws the bundled colour emoji over their text', async () => {
    const result = await convertDocx(makeDocx({ body: p('Done ✅ Team 👨‍👩‍👧') }), { compress: false });
    expect(result.warnings).toEqual([]);
    // The family's one glyph copies as all five characters (UTF-16 in the ToUnicode map).
    expect(pdfText(result.pdf)).toContain('<D83DDC68200DD83DDC69200DD83DDC67>');
    expect(pdfText(result.pdf)).toMatch(/\/BaseFont \/[A-Z]{6}\+NotoColorEmoji/);
    expect(pdfText(result.pdf)).toContain('/Subtype /Form');
    if (!ghostscript) return;
    // Ghostscript's text output garbles characters beyond U+FFFF, so it checks only ✅.
    expect(pdfPagesText(result.pdf)[0]).toContain('✅');
    // ✅ is a green box.
    const { rgb } = ghostscriptPixels(result.pdf, 72);
    let green = 0;
    for (let i = 0; i < rgb.length; i += 3) if (rgb[i + 1] > rgb[i] + 30 && rgb[i + 1] > rgb[i + 2] + 60) green++;
    expect(green).toBeGreaterThan(20);
  });
});
