import { fontFor } from '../fonts/script.js';
import { deobfuscateFont } from '../fonts/odttf.js';
import { attr, child, children, decodeEntities, descendant, resolveAlternateContent, textContent, type XmlElement } from '../util/xml.js';
import { parseDrawing, parseVml, loadImage, type DrawingContext } from './drawing.js';
import type {
  Block,
  Border,
  CellBorders,
  CellMargins,
  Columns,
  DocSettings,
  DocxDocument,
  HeaderFooterRefs,
  FontRequest,
  ImageRef,
  Inline,
  Link,
  Note,
  NumberingLabel,
  PageBorders,
  Paragraph,
  RunProps,
  Section,
  Table,
  TableCell,
  TableRow,
} from './model.js';
import { Numbering } from './numbering.js';
import type { ArchiveLimits } from '../util/zip.js';
import { OpcPackage } from './package.js';
import {
  finalizePara,
  finalizeRun,
  mergePara,
  mergeRun,
  mergeTbl,
  mergeTc,
  mergeTr,
  parseParaDelta,
  parseRunDelta,
  parseTblDelta,
  parseTcDelta,
  parseTrDelta,
  themeColorFn,
  type ParaDelta,
  type PropContext,
  type RunDelta,
  type TableBordersDelta,
  type TableLook,
  type TcDelta,
} from './props.js';
import { StyleSheet } from './styles.js';
import { DEFAULT_THEME, parseTheme, type Theme } from './theme.js';
import { emu, hexColor, num, onOff, parseBorder, twip, twipAttr } from './values.js';

interface TableContext {
  tableRun: RunDelta;
  tablePara: ParaDelta;
  override: boolean;
}

interface PartContext {
  part: string;
  table?: TableContext;
}

interface FieldFrame {
  instr: string;
  phase: 'instr' | 'result';
  kind?: 'page' | 'numpages' | 'sectionpages';
  format?: string;
  link?: Link;
  inline?: Inline & { type: 'field' };
  cached: string;
  /** A TOC field showing its result (the entries). */
  toc?: boolean;
  /** State of a legacy form checkbox field. */
  checkbox?: boolean;
}

const DEFAULT_SECTION_TWIPS = { w: 12240, h: 15840, margin: 1440, header: 720, footer: 720 };

export function parseDocx(data: Uint8Array, limits?: ArchiveLimits): DocxDocument {
  return new DocxParser(data, limits).parse();
}

class DocxParser {
  private pkg: OpcPackage;
  private main: string;
  private theme: Theme;
  private ctx: PropContext;
  private styles: StyleSheet;
  private numbering: Numbering;
  private settings: DocSettings;
  private fieldStack: FieldFrame[] = [];
  private images = new Map<string, ImageRef>();
  private fontRequests = new Map<string, FontRequest>();
  private warnings: string[] = [];
  private headerCache = new Map<string, Block[]>();
  private prevHeaders: HeaderFooterRefs = {};
  private prevFooters: HeaderFooterRefs = {};
  private hyperlinkStack: Link[] = [];

  constructor(data: Uint8Array, limits?: ArchiveLimits) {
    this.pkg = new OpcPackage(data, limits);
    this.main = this.pkg.mainDocumentPart();
    const themeRel = this.pkg.relByType(this.main, '/theme');
    this.theme = themeRel ? parseTheme(this.pkg.xml(themeRel.target)) : DEFAULT_THEME;
    const settingsRel = this.pkg.relByType(this.main, '/settings');
    const settingsXml = settingsRel ? this.pkg.xml(settingsRel.target) : undefined;
    const mapping = child(settingsXml, 'w:clrSchemeMapping');
    const colorMapping: Record<string, string> = {};
    if (mapping) {
      for (const [k, v] of Object.entries(mapping.attrs)) {
        const key = k.replace(/^w:/, '');
        const val = v === 'dark1' ? 'dk1' : v === 'light1' ? 'lt1' : v === 'dark2' ? 'dk2' : v === 'light2' ? 'lt2' : v;
        colorMapping[key] = val;
      }
    }
    this.ctx = { theme: this.theme, colorMapping };
    this.settings = this.parseSettings(settingsXml);
    const stylesRel = this.pkg.relByType(this.main, '/styles');
    this.styles = new StyleSheet(stylesRel ? this.pkg.xml(stylesRel.target) : undefined, this.ctx);
    const numRel = this.pkg.relByType(this.main, '/numbering');
    this.numbering = new Numbering(numRel ? this.pkg.xml(numRel.target) : undefined, this.ctx, (styleId) => this.styles.paraStyle(styleId).numId);
  }

  private parseSettings(s: XmlElement | undefined): DocSettings {
    const compat = child(s, 'w:compat');
    let mode = 12;
    let override = false;
    for (const cs of children(compat, 'w:compatSetting')) {
      if (cs.attrs['w:name'] === 'compatibilityMode') mode = num(cs.attrs['w:val']) ?? mode;
      if (cs.attrs['w:name'] === 'overrideTableStyleFontSizeAndJustification') override = cs.attrs['w:val'] === '1';
    }
    if (!compat || !children(compat, 'w:compatSetting').length) mode = compat ? 12 : 15;
    return {
      defaultTabStop: twipAttr(child(s, 'w:defaultTabStop'), 'w:val') ?? 36,
      evenAndOddHeaders: !!onOff(child(s, 'w:evenAndOddHeaders')),
      mirrorMargins: !!onOff(child(s, 'w:mirrorMargins')),
      compatibilityMode: mode,
      doNotExpandShiftReturn: !!onOff(child(compat, 'w:doNotExpandShiftReturn')),
      overrideTableStyleFontSizeAndJustification: override,
      autoHyphenation: !!onOff(child(s, 'w:autoHyphenation')),
      displayBackgroundShape: !!onOff(child(s, 'w:displayBackgroundShape')),
      doNotUseHTMLParagraphAutoSpacing: !!onOff(child(compat, 'w:doNotUseHTMLParagraphAutoSpacing')),
      footnotePosition: child(child(s, 'w:footnotePr'), 'w:pos')?.attrs['w:val'] ?? 'pageBottom',
      endnotePosition: child(child(s, 'w:endnotePr'), 'w:pos')?.attrs['w:val'] ?? 'docEnd',
    };
  }

  parse(): DocxDocument {
    const doc = this.pkg.xml(this.main);
    const body = child(doc, 'w:body');
    if (!body) throw new Error('Document has no body');
    const pctx: PartContext = { part: this.main };
    const blocks = this.parseBlocks(body, pctx);
    const finalSectPr = child(body, 'w:sectPr');
    const finalSection = this.parseSection(finalSectPr, pctx);

    const footnotes = new Map<string, Note>();
    const endnotes = new Map<string, Note>();
    let footnoteSeparator: Block[] | null = null;
    let footnoteContinuationSeparator: Block[] | null = null;
    let endnoteSeparator: Block[] | null = null;
    for (const [relType, elName, map] of [
      ['/footnotes', 'w:footnote', footnotes],
      ['/endnotes', 'w:endnote', endnotes],
    ] as const) {
      const rel = this.pkg.relByType(this.main, relType);
      if (!rel) continue;
      const root = this.pkg.xml(rel.target);
      for (const n of children(root, elName)) {
        const id = n.attrs['w:id'] ?? '';
        const type = n.attrs['w:type'];
        this.fieldStack = [];
        const blocks2 = this.parseBlocks(n, { part: rel.target });
        if (type === 'separator') {
          if (elName === 'w:footnote') footnoteSeparator = blocks2;
          else endnoteSeparator = blocks2;
        } else if (type === 'continuationSeparator') {
          if (elName === 'w:footnote') footnoteContinuationSeparator = blocks2;
        } else if (!type || type === 'normal') map.set(id, { id, blocks: blocks2 });
      }
    }

    const fontHints = new Map<string, { family?: string; pitch?: string; panose?: string; charset?: string; altName?: string }>();
    const embeddedFonts: DocxDocument['embeddedFonts'] = [];
    const ftRel = this.pkg.relByType(this.main, '/fontTable');
    if (ftRel) {
      const ft = this.pkg.xml(ftRel.target);
      for (const f of children(ft, 'w:font')) {
        const name = f.attrs['w:name'] ?? '';
        fontHints.set(name, {
          family: child(f, 'w:family')?.attrs['w:val'],
          pitch: child(f, 'w:pitch')?.attrs['w:val'],
          panose: child(f, 'w:panose1')?.attrs['w:val'],
          charset: child(f, 'w:charset')?.attrs['w:val'],
          altName: child(f, 'w:altName')?.attrs['w:val'],
        });
        for (const [tag, bold, italic] of [
          ['w:embedRegular', false, false],
          ['w:embedBold', true, false],
          ['w:embedItalic', false, true],
          ['w:embedBoldItalic', true, true],
        ] as const) {
          const e = child(f, tag);
          if (!e) continue;
          const rel = this.pkg.rels(ftRel.target).get(e.attrs['r:id'] ?? '');
          const bytes = rel ? this.pkg.bytes(rel.target) : null;
          if (!bytes) continue;
          const key = e.attrs['w:fontKey'];
          embeddedFonts.push({ family: name, bold, italic, data: key ? deobfuscateFont(bytes, key) : bytes });
        }
      }
    }

    const info: DocxDocument['info'] = {};
    const core = this.pkg.xml('docProps/core.xml');
    if (core) {
      const get = (n: string) => textContent(child(core, n)).trim() || undefined;
      info.title = get('dc:title');
      info.subject = get('dc:subject');
      info.creator = get('dc:creator');
      info.keywords = get('cp:keywords');
      info.description = get('dc:description');
    }

    const bg = child(doc, 'w:background');
    return {
      body: blocks,
      finalSection,
      settings: this.settings,
      footnotes,
      endnotes,
      footnoteSeparator,
      footnoteContinuationSeparator,
      endnoteSeparator,
      info,
      background: bg ? (hexColor(bg.attrs['w:color']) ?? null) : null,
      fontRequests: [...this.fontRequests.values()],
      embeddedFonts,
      fontHints,
      warnings: this.warnings,
    };
  }

  // -------------------------------------------------------------------------------------------
  // Blocks

  parseBlocks(container: XmlElement, pctx: PartContext): Block[] {
    const out: Block[] = [];
    this.collectBlocks(container, pctx, out);
    // Word does not lay out an empty paragraph whose mark is hidden (it joins the next paragraph).
    // The last paragraph of the body, a cell or a note stays: it closes its container.
    return out.filter(
      (b, i) => i === out.length - 1 || b.type !== 'paragraph' || !b.markProps.hidden || b.sectPr || !b.content.every((c) => c.type === 'renderedPageBreak'),
    );
  }

  private importChunk(el: XmlElement, pctx: PartContext): Block[] {
    const id = el.attrs['r:id'] ?? '';
    const rel = this.pkg.rels(pctx.part).get(id);
    const type = rel && !rel.external ? this.pkg.contentType(rel.target) : undefined;
    if (!rel || (type !== 'text/html' && type !== 'application/xhtml+xml' && type !== 'text/plain')) {
      this.warnings.push('Skipped unsupported or missing altChunk import: ' + (rel?.target ?? id));
      return [];
    }
    let text = this.pkg.zip.readText(rel.target);
    if (text === null) {
      this.warnings.push('Missing altChunk part: ' + rel.target);
      return [];
    }
    if (type !== 'text/plain') {
      // Import readable text only. HTML is never attached to a browser DOM or executed.
      this.warnings.push('Imported HTML altChunk as plain text: ' + rel.target);
      text = text.replace(/<!--[\s\S]*?-->/g, '')
        .replace(/<(head|script|style)\b[^>]*>[\s\S]*?<\/\1\s*>/gi, '')
        .replace(/\s+/g, ' ')
        .replace(/<\/?(?:p|div|li|h[1-6]|br|tr|table|section|article)\b[^>]*>/gi, '\n')
        .replace(/<\/t[dh]\s*>/gi, ' ')
        .replace(/<\/?[a-z!](?:[^>"']|"[^"]*"|'[^']*')*>/gi, '');
      text = decodeEntities(text.replace(/&nbsp;/gi, '\u00a0'));
    }
    const lines = text.split(/\r\n?|\n/);
    const blocks: Block[] = [];
    for (const line of lines) {
      if (type !== 'text/plain' && !line.trim()) continue;
      const para = this.parseParagraph({ name: 'w:p', attrs: {}, children: [] }, pctx);
      this.pushText(para, type === 'text/plain' ? line : line.trim(), para.markProps);
      blocks.push(para);
    }
    return blocks;
  }

  private collectBlocks(container: XmlElement, pctx: PartContext, out: Block[]): void {
    for (const el of children(container)) {
      switch (el.name) {
        case 'w:p':
          out.push(this.parseParagraph(el, pctx));
          break;
        case 'w:altChunk':
          out.push(...this.importChunk(el, pctx));
          break;
        case 'w:tbl':
          out.push(this.parseTable(el, pctx));
          break;
        case 'w:sdt':
          this.collectBlocks(child(el, 'w:sdtContent') ?? el, pctx, out);
          break;
        case 'w:customXml':
        case 'w:ins':
        case 'w:moveTo':
        case 'w:smartTag':
          this.collectBlocks(el, pctx, out);
          break;
        case 'w:del':
        case 'w:moveFrom':
          break;
        case 'mc:AlternateContent': {
          const alternatives = resolveAlternateContent(el);
          if (!alternatives.length) this.warnings.push('Skipped alternate block without a supported choice or fallback in ' + pctx.part);
          for (const alt of alternatives) this.collectBlocks({ name: 'x', attrs: {}, children: [alt] }, pctx, out);
          break;
        }
        default:
          if (descendant(el, 'w:t') || descendant(el, 'w:drawing') || descendant(el, 'w:pict')) {
            this.warnings.push('Skipped unsupported content block ' + el.name + ' in ' + pctx.part);
          }
      }
    }
  }

  // -------------------------------------------------------------------------------------------
  // Paragraphs

  private parseParagraph(p: XmlElement, pctx: PartContext): Paragraph {
    const pPr = child(p, 'w:pPr');
    const direct = parseParaDelta(pPr, this.ctx);
    const styleId = direct.styleId ?? this.styles.defaultParagraphStyle;
    const isDefaultStyle = styleId === this.styles.defaultParagraphStyle;
    const styleP = this.styles.paraStyle(styleId);
    const numId = direct.numId ?? styleP.numId;
    const fromStyle = direct.numId === undefined && styleP.numId !== undefined;
    const ilvl = direct.ilvl ?? styleP.ilvl ?? 0;
    const hasNumbering = this.numbering.exists(numId);
    const lvl = hasNumbering ? this.numbering.level(numId!, ilvl) : undefined;

    let merged: ParaDelta = this.styles.defaultsPara;
    const tableP = pctx.table?.tablePara;
    if (tableP) merged = mergePara(merged, tableP);
    if (lvl && fromStyle) merged = mergePara(mergePara(merged, lvl.pPr), styleP);
    else if (lvl) merged = mergePara(mergePara(merged, styleP), lvl.pPr);
    else merged = mergePara(merged, styleP);
    if (tableP?.align !== undefined && pctx.table?.override && isDefaultStyle) merged = { ...merged, align: tableP.align };
    const directNoStyle = { ...direct };
    delete directNoStyle.styleId;
    merged = mergePara(merged, directNoStyle);
    const props = finalizePara(merged);
    props.styleId = styleId ?? '';

    const tableOverridesSize = !!pctx.table?.override && isDefaultStyle;
    const markDirect = parseRunDelta(child(pPr, 'w:rPr'), this.ctx);
    const markDelta = this.styles.runDelta({ paraStyle: styleId, charStyle: markDirect.styleId, tableRun: pctx.table?.tableRun, direct: markDirect, tableOverridesSize });
    const markProps = finalizeRun(markDelta);
    this.noteFonts(markProps, '');

    const para: Paragraph = { type: 'paragraph', props, markProps, content: [], numbering: null };

    if (lvl && hasNumbering) {
      const text = this.numbering.next(numId!, ilvl);
      if (text !== null) {
        const labelProps = finalizeRun(mergeRun(markDelta, lvl.rPr));
        // Word never underlines or highlights the number from the paragraph mark alone.
        if (!lvl.rPr.underline) labelProps.underline = null;
        labelProps.highlight = null;
        let picture: ImageRef | undefined;
        if (lvl.picBulletId !== undefined) {
          const rid = this.numbering.picBullets.get(lvl.picBulletId);
          const numRel = this.pkg.relByType(this.main, '/numbering');
          if (rid && numRel) picture = loadImage(this.drawingCtx({ part: numRel.target }), rid);
        }
        const label: NumberingLabel = { text, props: labelProps, suffix: lvl.suffix, align: lvl.jc, picture };
        para.numbering = label;
        this.noteFonts(labelProps, text);
      }
    }

    const sectPr = child(pPr, 'w:sectPr');
    this.parseInlines(p, para, pctx, styleId, tableOverridesSize);
    if (sectPr && pctx.part === this.main) para.sectPr = this.parseSection(sectPr, pctx);
    return para;
  }

  private parseInlines(container: XmlElement, para: Paragraph, pctx: PartContext, styleId: string | undefined, tableOverridesSize: boolean): void {
    for (const el of children(container)) {
      switch (el.name) {
        case 'w:r':
          this.parseRun(el, para, pctx, styleId, tableOverridesSize);
          break;
        case 'w:hyperlink': {
          const rid = el.attrs['r:id'];
          const link: Link = {};
          if (rid) {
            const rel = this.pkg.rels(pctx.part).get(rid);
            if (rel) link.url = rel.target;
          }
          if (el.attrs['w:anchor']) link.anchor = el.attrs['w:anchor'];
          this.hyperlinkStack.push(link);
          this.parseInlines(el, para, pctx, styleId, tableOverridesSize);
          this.hyperlinkStack.pop();
          break;
        }
        case 'w:fldSimple': {
          const instr = el.attrs['w:instr'] ?? '';
          const info = analyzeField(instr);
          if (info.kind) {
            const firstRun = child(el, 'w:r');
            const props = this.runProps(child(firstRun, 'w:rPr'), styleId, pctx, tableOverridesSize);
            if (!this.suppressed())
              para.content.push({ type: 'field', field: info.kind, format: info.format, props, cached: textContent(el) });
          } else {
            if (info.link) this.hyperlinkStack.push(info.link);
            this.parseInlines(el, para, pctx, styleId, tableOverridesSize);
            if (info.link) this.hyperlinkStack.pop();
          }
          break;
        }
        case 'w:ins':
        case 'w:moveTo':
        case 'w:smartTag':
        case 'w:customXml':
        case 'w:dir':
        case 'w:bdo':
          this.parseInlines(el, para, pctx, styleId, tableOverridesSize);
          break;
        case 'w:sdt':
          this.parseInlines(child(el, 'w:sdtContent') ?? el, para, pctx, styleId, tableOverridesSize);
          break;
        case 'w:bookmarkStart':
          if (el.attrs['w:name']) para.content.push({ type: 'bookmark', name: el.attrs['w:name'] });
          break;
        case 'm:oMathPara':
        case 'm:oMath':
          this.parseMath(el, para, pctx, styleId, tableOverridesSize);
          break;
        case 'mc:AlternateContent':
          this.parseInlines({ name: 'x', attrs: {}, children: resolveAlternateContent(el) }, para, pctx, styleId, tableOverridesSize);
          break;
      }
    }
  }

  /** Office Math is rendered linearly (base text of each math run) in the run's font. */
  private parseMath(el: XmlElement, para: Paragraph, pctx: PartContext, styleId: string | undefined, tableOverridesSize: boolean): void {
    for (const c of children(el)) {
      if (c.name === 'm:r') {
        const props = this.runProps(child(c, 'w:rPr'), styleId, pctx, tableOverridesSize);
        const text = children(c, 'm:t')
          .map((t) => textContent(t))
          .join('');
        if (text && !this.suppressed()) {
          if (!child(c, 'w:rPr')?.children.some((x) => typeof x !== 'string' && x.name === 'w:rFonts')) {
            props.fontAscii = props.fontHAnsi = 'Cambria Math';
          }
          props.italic = !child(child(c, 'm:rPr'), 'm:sty') && /[A-Za-z]/.test(text);
          this.pushText(para, text, props);
        }
      } else if (c.name === 'm:f') {
        // Fraction: num/den
        this.parseMath(child(c, 'm:num') ?? c, para, pctx, styleId, tableOverridesSize);
        const props = this.runProps(undefined, styleId, pctx, tableOverridesSize);
        this.pushText(para, '/', props);
        this.parseMath(child(c, 'm:den') ?? c, para, pctx, styleId, tableOverridesSize);
      } else if (c.name === 'w:r') {
        this.parseRun(c, para, pctx, styleId, tableOverridesSize);
      } else if (!c.name.endsWith('Pr')) {
        this.parseMath(c, para, pctx, styleId, tableOverridesSize);
      }
    }
  }

  private runProps(rPr: XmlElement | undefined, styleId: string | undefined, pctx: PartContext, tableOverridesSize: boolean): RunProps {
    const direct = parseRunDelta(rPr, this.ctx);
    let charStyle = direct.styleId;
    // Word shows table-of-contents entries in their TOC paragraph styles: the Hyperlink character
    // style that the TOC field's \h switch puts on every entry is not drawn (the link still works).
    if (charStyle && this.styles.styles.get(charStyle)?.name.toLowerCase() === 'hyperlink' && this.fieldStack.some((f) => f.toc)) charStyle = undefined;
    const delta = this.styles.runDelta({ paraStyle: styleId, charStyle, tableRun: pctx.table?.tableRun, direct, tableOverridesSize });
    const props = finalizeRun(delta);
    props.styleId = direct.styleId;
    return props;
  }

  private suppressed(): boolean {
    return this.fieldStack.some((f) => f.phase === 'instr' || (f.kind !== undefined && f.phase === 'result'));
  }

  private currentLink(): Link | undefined {
    for (let i = this.fieldStack.length - 1; i >= 0; i--) if (this.fieldStack[i].link) return this.fieldStack[i].link;
    return this.hyperlinkStack[this.hyperlinkStack.length - 1];
  }

  private noteFonts(props: RunProps, text: string): void {
    for (const ch of text || ' ') {
      const cp = ch.codePointAt(0)!;
      const { family, bold, italic } = fontFor(cp, props);
      const key = family + '\0' + Number(bold) + Number(italic);
      let request = this.fontRequests.get(key);
      if (!request) {
        request = { family, bold, italic, codePoints: new Set() };
        this.fontRequests.set(key, request);
      }
      request.codePoints.add(cp);
    }
  }

  private pushText(para: Paragraph, text: string, props: RunProps): void {
    if (!text) return;
    const frame = this.fieldStack[this.fieldStack.length - 1];
    if (frame && frame.kind && frame.phase === 'result') frame.cached += text;
    if (this.suppressed() || props.hidden) return;
    this.noteFonts(props, text);
    const link = this.currentLink();
    const last = para.content[para.content.length - 1];
    if (last && last.type === 'text' && last.props === props && last.link === link) {
      last.text += text;
      return;
    }
    para.content.push({ type: 'text', text, props, link });
  }

  private parseRun(r: XmlElement, para: Paragraph, pctx: PartContext, styleId: string | undefined, tableOverridesSize: boolean): void {
    const props = this.runProps(child(r, 'w:rPr'), styleId, pctx, tableOverridesSize);
    // Breaks, tabs and pictures are measured with the run's font too, so load it even when the
    // run has no text.
    const visible = () => !this.suppressed() && !props.hidden;
    if (visible()) this.noteFonts(props, '');
    const items = children(r);
    for (let idx = 0; idx < items.length; idx++) {
      const el = items[idx];
      switch (el.name) {
        case 'w:t': {
          let t = textContent(el);
          if (el.attrs['xml:space'] !== 'preserve') t = t.replace(/^[ \t\r\n]+|[ \t\r\n]+$/g, '');
          t = t.replace(/[\r\n]/g, '');
          this.pushText(para, t, props);
          break;
        }
        case 'w:tab':
          if (visible()) para.content.push({ type: 'tab', props, link: this.currentLink() });
          break;
        case 'w:ptab': {
          if (!visible()) break;
          const a = el.attrs['w:alignment'];
          const leader = el.attrs['w:leader'];
          para.content.push({
            type: 'ptab',
            props,
            alignment: a === 'center' ? 'center' : a === 'right' ? 'right' : 'left',
            relativeTo: el.attrs['w:relativeTo'] === 'indent' ? 'indent' : 'margin',
            leader: leader === 'dot' || leader === 'hyphen' || leader === 'underscore' || leader === 'middleDot' ? leader : 'none',
          });
          break;
        }
        case 'w:br': {
          if (!visible()) break;
          const type = el.attrs['w:type'];
          para.content.push({ type: 'break', breakType: type === 'page' ? 'page' : type === 'column' ? 'column' : 'line', clear: el.attrs['w:clear'], props });
          break;
        }
        case 'w:cr':
          if (visible()) para.content.push({ type: 'break', breakType: 'line', props });
          break;
        case 'w:softHyphen':
          this.pushText(para, '\u00ad', props);
          break;
        case 'w:noBreakHyphen':
          this.pushText(para, '\u2011', props);
          break;
        case 'w:sym': {
          const font = el.attrs['w:font'];
          const code = parseInt(el.attrs['w:char'] ?? '0', 16);
          if (!code) break;
          const symProps: RunProps = font ? { ...props, fontAscii: font, fontHAnsi: font, fontEastAsia: font, fontCs: font } : props;
          this.pushText(para, String.fromCodePoint(code), symProps);
          break;
        }
        case 'w:drawing': {
          if (!visible()) break;
          const d = parseDrawing(el, this.drawingCtx(pctx));
          if (d && !d.hidden) para.content.push({ type: 'drawing', drawing: d, props, link: this.currentLink() });
          break;
        }
        case 'w:pict':
        case 'w:object': {
          if (!visible()) break;
          const d = parseVml(el, this.drawingCtx(pctx));
          if (d && !d.hidden && d.kind !== 'unsupported') para.content.push({ type: 'drawing', drawing: d, props, link: this.currentLink() });
          break;
        }
        case 'mc:AlternateContent': {
          const alt = resolveAlternateContent(el);
          const fake: XmlElement = { name: 'w:r', attrs: {}, children: [child(r, 'w:rPr') ?? { name: 'w:rPr', attrs: {}, children: [] }, ...alt] };
          this.parseRun(fake, para, pctx, styleId, tableOverridesSize);
          break;
        }
        case 'w:fldChar': {
          const type = el.attrs['w:fldCharType'];
          if (type === 'begin') {
            // Legacy form checkboxes (FORMCHECKBOX) keep their state in the field's form data.
            const box = child(child(el, 'w:ffData'), 'w:checkBox');
            const state = child(box, 'w:checked') ?? child(box, 'w:default');
            const checked = state ? !['0', 'false', 'off'].includes(state.attrs['w:val'] ?? '1') : false;
            this.fieldStack.push({ instr: '', phase: 'instr', cached: '', checkbox: box ? checked : undefined });
          }
          else if (type === 'separate') {
            const f = this.fieldStack[this.fieldStack.length - 1];
            if (f && f.phase === 'instr') {
              const info = analyzeField(f.instr);
              f.phase = 'result';
              f.link = info.link;
              f.toc = /^TOC\b/i.test(f.instr.trim());
              if (info.kind) {
                f.kind = info.kind;
                f.format = info.format;
                // Only emit when every enclosing field is showing its result.
                const outerVisible = this.fieldStack.slice(0, -1).every((o) => o.phase === 'result' && !o.kind);
                if (outerVisible && !props.hidden) {
                  const inline: Inline & { type: 'field' } = { type: 'field', field: info.kind, format: info.format, props, cached: '' };
                  f.inline = inline;
                  para.content.push(inline);
                }
              }
            }
          } else if (type === 'end') {
            const f = this.fieldStack.pop();
            if (f?.checkbox !== undefined && !this.suppressed()) this.pushText(para, f.checkbox ? '\u2612' : '\u2610', props);
            if (f) {
              if (f.phase === 'instr') {
                // Field without a result: dynamic fields still render.
                const info = analyzeField(f.instr);
                if (info.kind && !this.suppressed() && !props.hidden)
                  para.content.push({ type: 'field', field: info.kind, format: info.format, props, cached: '' });
              } else if (f.inline) f.inline.cached = f.cached;
            }
          }
          break;
        }
        case 'w:instrText': {
          const f = this.fieldStack[this.fieldStack.length - 1];
          if (f && f.phase === 'instr') f.instr += textContent(el);
          break;
        }
        case 'w:footnoteReference':
        case 'w:endnoteReference':
          if (visible())
            para.content.push({
              type: 'noteRef',
              noteType: el.name === 'w:footnoteReference' ? 'footnote' : 'endnote',
              id: el.attrs['w:id'] ?? '',
              props,
              customMark: el.attrs['w:customMarkFollows'] === '1' || el.attrs['w:customMarkFollows'] === 'true' ? '' : undefined,
            });
          break;
        case 'w:footnoteRef':
        case 'w:endnoteRef':
          if (visible()) para.content.push({ type: 'noteNumber', props });
          break;
        case 'w:separator':
        case 'w:continuationSeparator':
          para.content.push({ type: 'separator', continuation: el.name === 'w:continuationSeparator', props });
          break;
        case 'w:lastRenderedPageBreak':
          para.content.push({ type: 'renderedPageBreak' });
          break;
        case 'w:ruby': {
          const base = child(el, 'w:rubyBase');
          if (base) this.parseInlines(base, para, pctx, styleId, tableOverridesSize);
          break;
        }
      }
    }
  }

  private drawingCtx(pctx: { part: string }): DrawingContext {
    return {
      pkg: this.pkg,
      part: pctx.part,
      theme: this.theme,
      colorMapping: this.ctx.colorMapping,
      images: this.images,
      parseBlocks: (container) => {
        // Text box content is a separate story: field state must not leak in or out.
        const saved = this.fieldStack;
        const savedLinks = this.hyperlinkStack;
        this.fieldStack = [];
        this.hyperlinkStack = [];
        const blocks = this.parseBlocks(container, { part: pctx.part });
        this.fieldStack = saved;
        this.hyperlinkStack = savedLinks;
        return blocks;
      },
    };
  }

  // -------------------------------------------------------------------------------------------
  // Tables

  private parseTable(tbl: XmlElement, pctx: PartContext): Table {
    const direct = parseTblDelta(child(tbl, 'w:tblPr'), this.ctx);
    const styleId = direct.styleId ?? this.styles.defaultTableStyle;
    const ts = this.styles.tableStyle(styleId);
    const tblD = mergeTbl(ts.tbl, direct);
    const look: TableLook = direct.look ?? tblD.look ?? { firstRow: true, lastRow: false, firstColumn: true, lastColumn: false, noHBand: false, noVBand: true };
    const grid = children(child(tbl, 'w:tblGrid'), 'w:gridCol').map((g) => twipAttr(g, 'w:w') ?? 0);
    const tableBorders: TableBordersDelta = tblD.borders ?? {};
    const margins: CellMargins = { top: 0, bottom: 0, left: 5.4, right: 5.4, ...(tblD.cellMargins ?? {}) };

    // Gather rows/cells, unwrapping content controls and custom XML.
    const rowEls: XmlElement[] = [];
    const collectRows = (el: XmlElement) => {
      for (const c of children(el)) {
        if (c.name === 'w:tr') {
          if (!child(child(c, 'w:trPr'), 'w:del')) rowEls.push(c);
        }
        else if (c.name === 'w:sdt') collectRows(child(c, 'w:sdtContent') ?? c);
        else if (c.name === 'w:customXml' || c.name === 'w:ins') collectRows(c);
      }
    };
    collectRows(tbl);
    const cellEls = (tr: XmlElement): XmlElement[] => {
      const out: XmlElement[] = [];
      const walk = (el: XmlElement) => {
        for (const c of children(el)) {
          if (c.name === 'w:tc') out.push(c);
          else if (c.name === 'w:sdt') walk(child(c, 'w:sdtContent') ?? c);
          else if (c.name === 'w:customXml') walk(c);
        }
      };
      walk(tr);
      return out;
    };

    const nRows = rowEls.length;
    const nCols = Math.max(grid.length, 1);
    const rowBand = tblD.rowBandSize || 1;
    const colBand = tblD.colBandSize || 1;
    const rows: TableRow[] = [];
    for (let r = 0; r < nRows; r++) {
      const tr = rowEls[r];
      const trD = mergeTr(ts.tr, parseTrDelta(child(tr, 'w:trPr')));
      const row: TableRow = {
        height: trD.height ?? 0,
        heightRule: trD.heightRule ?? 'atLeast',
        cantSplit: !!trD.cantSplit,
        isHeader: !!trD.isHeader,
        gridBefore: trD.gridBefore ?? 0,
        gridAfter: trD.gridAfter ?? 0,
        cells: [],
      };
      // Row-level exceptions to the table properties (w:tblPrEx) override borders, shading and
      // cell margins for this row's cells.
      const ex = parseTblDelta(child(tr, 'w:tblPrEx'), this.ctx);
      const rowBorders: TableBordersDelta = ex.borders ? { ...tableBorders, ...ex.borders } : tableBorders;
      const rowShading = ex.shading !== undefined ? ex.shading : tblD.shading;
      const rowMargins: CellMargins = ex.cellMargins ? { ...margins, ...ex.cellMargins } : margins;
      let col = row.gridBefore;
      const tcs = cellEls(tr);
      for (const tc of tcs) {
        const tcDirect = parseTcDelta(child(tc, 'w:tcPr'), this.ctx);
        const span = Math.max(1, tcDirect.gridSpan ?? 1);
        const startCol = col;
        const endCol = col + span;
        col = endCol;

        const conds: Array<{ type: string; region: 'row' | 'col' | 'corner' | 'whole' }> = [{ type: 'wholeTable', region: 'whole' }];
        const isFirstRow = r === 0 && look.firstRow;
        const isLastRow = r === nRows - 1 && look.lastRow && nRows > 1;
        const isFirstCol = startCol === 0 && look.firstColumn;
        const isLastCol = endCol >= nCols && look.lastColumn;
        if (!look.noVBand) {
          const idx = startCol - (look.firstColumn ? 1 : 0);
          if (idx >= 0 && !isFirstCol && !isLastCol) conds.push({ type: Math.floor(idx / colBand) % 2 === 0 ? 'band1Vert' : 'band2Vert', region: 'col' });
        }
        if (!look.noHBand) {
          const idx = r - (look.firstRow ? 1 : 0);
          if (idx >= 0 && !isFirstRow && !isLastRow) conds.push({ type: Math.floor(idx / rowBand) % 2 === 0 ? 'band1Horz' : 'band2Horz', region: 'row' });
        }
        if (isFirstCol) conds.push({ type: 'firstCol', region: 'col' });
        if (isLastCol) conds.push({ type: 'lastCol', region: 'col' });
        if (isFirstRow) conds.push({ type: 'firstRow', region: 'row' });
        if (isLastRow) conds.push({ type: 'lastRow', region: 'row' });
        if (isFirstRow && isFirstCol) conds.push({ type: 'nwCell', region: 'corner' });
        if (isFirstRow && isLastCol) conds.push({ type: 'neCell', region: 'corner' });
        if (isLastRow && isFirstCol) conds.push({ type: 'swCell', region: 'corner' });
        if (isLastRow && isLastCol) conds.push({ type: 'seCell', region: 'corner' });

        let pD: ParaDelta = ts.pPr;
        let rD: RunDelta = ts.rPr;
        let tcD: TcDelta = ts.tc;
        const isTop = r === 0;
        const isBottom = r === nRows - 1;
        const isLeft = startCol === 0;
        const isRight = endCol >= nCols;
        const pick = (b: Border | null | undefined, fallback: Border | null | undefined) => (b !== undefined ? b : fallback);
        let top = isTop ? rowBorders.top : rowBorders.insideH;
        let bottom = isBottom ? rowBorders.bottom : rowBorders.insideH;
        let left = isLeft ? rowBorders.left : rowBorders.insideV;
        let right = isRight ? rowBorders.right : rowBorders.insideV;
        let shading = rowShading;
        for (const c of conds) {
          const cf = ts.cond.get(c.type);
          if (!cf) continue;
          pD = mergePara(pD, cf.pPr);
          rD = mergeRun(rD, cf.rPr);
          tcD = mergeTc(tcD, { ...cf.tc, borders: undefined });
          if (cf.tbl.shading !== undefined) shading = cf.tbl.shading;
          const cb = { ...(cf.tbl.borders ?? {}), ...(cf.tc.borders ?? {}) };
          if (c.region === 'row') {
            top = pick(cb.top, top);
            bottom = pick(cb.bottom, bottom);
            left = isLeft ? pick(cb.left, left) : pick(cb.insideV, left);
            right = isRight ? pick(cb.right, right) : pick(cb.insideV, right);
          } else if (c.region === 'col') {
            left = pick(cb.left, left);
            right = pick(cb.right, right);
            top = isTop ? pick(cb.top, top) : pick(cb.insideH, top);
            bottom = isBottom ? pick(cb.bottom, bottom) : pick(cb.insideH, bottom);
          } else {
            top = isTop ? pick(cb.top, top) : pick(cb.insideH, top);
            bottom = isBottom ? pick(cb.bottom, bottom) : pick(cb.insideH, bottom);
            left = isLeft ? pick(cb.left, left) : pick(cb.insideV, left);
            right = isRight ? pick(cb.right, right) : pick(cb.insideV, right);
          }
        }
        tcD = mergeTc(tcD, tcDirect);
        const db = tcDirect.borders ?? {};
        top = pick(db.top, top);
        bottom = pick(db.bottom, bottom);
        left = pick(db.left, left);
        right = pick(db.right, right);
        const borders: CellBorders = {
          top: top ?? null,
          bottom: bottom ?? null,
          left: left ?? null,
          right: right ?? null,
          tl2br: db.tl2br ?? null,
          tr2bl: db.tr2bl ?? null,
        };
        const tableCtx: TableContext = { tableRun: rD, tablePara: pD, override: this.settings.overrideTableStyleFontSizeAndJustification };
        const cell: TableCell = {
          gridSpan: span,
          vMerge: tcDirect.vMerge ?? null,
          width: tcDirect.width ?? { type: 'auto', value: 0 },
          borders,
          shading: tcD.shading !== undefined ? tcD.shading : (shading ?? null),
          margins: { ...rowMargins, ...(tcD.margins ?? {}) },
          vAlign: tcD.vAlign ?? 'top',
          noWrap: !!tcD.noWrap,
          hideMark: !!tcD.hideMark,
          textDirection: tcD.textDirection ?? 'lrTb',
          content: this.parseBlocks(tc, { part: pctx.part, table: tableCtx }),
        };
        row.cells.push(cell);
      }
      rows.push(row);
    }

    // Grid fallback: derive columns from cell widths when tblGrid is missing or short.
    let finalGrid = grid;
    const maxCols = Math.max(0, ...rows.map((r) => r.gridBefore + r.cells.reduce((s, c) => s + c.gridSpan, 0) + r.gridAfter));
    if (finalGrid.length < maxCols || finalGrid.every((w) => w <= 0)) {
      finalGrid = new Array(maxCols).fill(0);
      for (const row of rows) {
        let c = row.gridBefore;
        for (const cell of row.cells) {
          const w = cell.width.type === 'dxa' ? cell.width.value / cell.gridSpan : 0;
          for (let k = 0; k < cell.gridSpan; k++) if (w > 0 && !finalGrid[c + k]) finalGrid[c + k] = w;
          c += cell.gridSpan;
        }
      }
    }

    return {
      type: 'table',
      props: {
        styleId,
        width: tblD.width ?? { type: 'auto', value: 0 },
        align: tblD.align ?? 'left',
        indent: tblD.indent ?? 0,
        layout: tblD.layout ?? 'autofit',
        cellMargins: margins,
        cellSpacing: tblD.cellSpacing ?? 0,
        floating: direct.floating ? parseFloating(direct.floating) : null,
        bidiVisual: !!tblD.bidiVisual,
      },
      grid: finalGrid,
      rows,
    };
  }

  // -------------------------------------------------------------------------------------------
  // Sections, headers and footers

  private parseSection(sectPr: XmlElement | undefined, pctx: PartContext): Section {
    const pgSz = child(sectPr, 'w:pgSz');
    const pgMar = child(sectPr, 'w:pgMar');
    const t = (el: XmlElement | undefined, n: string, d: number) => twip(num(attr(el, n)) ?? d);
    let pageWidth = t(pgSz, 'w:w', DEFAULT_SECTION_TWIPS.w);
    let pageHeight = t(pgSz, 'w:h', DEFAULT_SECTION_TWIPS.h);
    if (attr(pgSz, 'w:orient') === 'landscape' && pageWidth < pageHeight) [pageWidth, pageHeight] = [pageHeight, pageWidth];
    const topRaw = num(attr(pgMar, 'w:top')) ?? DEFAULT_SECTION_TWIPS.margin;
    const bottomRaw = num(attr(pgMar, 'w:bottom')) ?? DEFAULT_SECTION_TWIPS.margin;

    const cols = child(sectPr, 'w:cols');
    const colEls = children(cols, 'w:col');
    const columns: Columns = {
      count: Math.max(1, num(attr(cols, 'w:num')) ?? (colEls.length || 1)),
      space: twip(num(attr(cols, 'w:space')) ?? 720),
      equalWidth: attr(cols, 'w:equalWidth') !== '0' && attr(cols, 'w:equalWidth') !== 'false',
      separator: attr(cols, 'w:sep') === '1' || attr(cols, 'w:sep') === 'true',
      cols: colEls.map((c) => ({ width: twip(num(c.attrs['w:w']) ?? 0), space: twip(num(c.attrs['w:space']) ?? 0) })),
    };
    if (colEls.length && !columns.equalWidth) columns.count = colEls.length;

    const headers: HeaderFooterRefs = { ...this.prevHeaders };
    const footers: HeaderFooterRefs = { ...this.prevFooters };
    for (const ref of children(sectPr, 'w:headerReference').concat(children(sectPr, 'w:footerReference'))) {
      const type = (ref.attrs['w:type'] ?? 'default') as keyof HeaderFooterRefs;
      const rel = this.pkg.rels(pctx.part).get(ref.attrs['r:id'] ?? '');
      if (!rel) continue;
      const blocks = this.headerFooter(rel.target, ref.name === 'w:headerReference' ? 'w:hdr' : 'w:ftr');
      if (ref.name === 'w:headerReference') headers[type] = blocks;
      else footers[type] = blocks;
    }
    this.prevHeaders = headers;
    this.prevFooters = footers;

    const typeVal = attr(child(sectPr, 'w:type'), 'w:val');
    const pgNum = child(sectPr, 'w:pgNumType');
    const grid = child(sectPr, 'w:docGrid');
    const fnPr = child(sectPr, 'w:footnotePr');
    return {
      pageWidth,
      pageHeight,
      marginTop: twip(Math.abs(topRaw)),
      marginBottom: twip(Math.abs(bottomRaw)),
      marginTopExact: topRaw < 0,
      marginBottomExact: bottomRaw < 0,
      marginLeft: t(pgMar, attr(pgMar, 'w:left') === undefined && attr(pgMar, 'w:start') !== undefined ? 'w:start' : 'w:left', DEFAULT_SECTION_TWIPS.margin),
      marginRight: t(pgMar, attr(pgMar, 'w:right') === undefined && attr(pgMar, 'w:end') !== undefined ? 'w:end' : 'w:right', DEFAULT_SECTION_TWIPS.margin),
      header: t(pgMar, 'w:header', DEFAULT_SECTION_TWIPS.header),
      footer: t(pgMar, 'w:footer', DEFAULT_SECTION_TWIPS.footer),
      gutter: t(pgMar, 'w:gutter', 0),
      columns,
      type: (['nextPage', 'continuous', 'evenPage', 'oddPage', 'nextColumn'].includes(typeVal ?? '') ? typeVal : 'nextPage') as Section['type'],
      titlePage: !!onOff(child(sectPr, 'w:titlePg')),
      headers,
      footers,
      pageNumberStart: num(attr(pgNum, 'w:start')) ?? null,
      pageNumberFormat: attr(pgNum, 'w:fmt') ?? 'decimal',
      vAlign: (attr(child(sectPr, 'w:vAlign'), 'w:val') ?? 'top') as Section['vAlign'],
      gridType: attr(grid, 'w:type') ?? 'default',
      gridLinePitch: twip(num(attr(grid, 'w:linePitch')) ?? 360),
      pageBorders: this.parsePageBorders(child(sectPr, 'w:pgBorders')),
      footnoteRestart: (attr(child(fnPr, 'w:numRestart'), 'w:val') ?? 'continuous') as Section['footnoteRestart'],
      footnoteStart: num(attr(child(fnPr, 'w:numStart'), 'w:val')) ?? 1,
      footnoteFormat: attr(child(fnPr, 'w:numFmt'), 'w:val') ?? 'decimal',
    };
  }

  private parsePageBorders(el: XmlElement | undefined): PageBorders | null {
    if (!el) return null;
    const tc = themeColorFn(this.ctx);
    const side = (n: string) => parseBorder(child(el, n), tc) ?? null;
    const pb: PageBorders = {
      top: side('w:top'),
      bottom: side('w:bottom'),
      left: side('w:left'),
      right: side('w:right'),
      offsetFrom: el.attrs['w:offsetFrom'] === 'page' ? 'page' : 'text',
      display: (el.attrs['w:display'] ?? 'allPages') as PageBorders['display'],
      zOrder: el.attrs['w:zOrder'] === 'back' ? 'back' : 'front',
    };
    if (!pb.top && !pb.bottom && !pb.left && !pb.right) return null;
    return pb;
  }

  private headerFooter(part: string, rootName: string): Block[] {
    const cached = this.headerCache.get(part);
    if (cached) return cached;
    const root = this.pkg.xml(part);
    const savedFields = this.fieldStack;
    this.fieldStack = [];
    const blocks = root && root.name === rootName ? this.parseBlocks(root, { part }) : root ? this.parseBlocks(root, { part }) : [];
    this.fieldStack = savedFields;
    this.headerCache.set(part, blocks);
    return blocks;
  }
}

function parseFloating(el: XmlElement) {
  const t = (n: string) => twip(num(el.attrs[n]) ?? 0);
  return {
    horzAnchor: el.attrs['w:horzAnchor'] ?? 'text',
    vertAnchor: el.attrs['w:vertAnchor'] ?? 'margin',
    x: t('w:tblpX'),
    y: t('w:tblpY'),
    xAlign: el.attrs['w:tblpXSpec'],
    yAlign: el.attrs['w:tblpYSpec'],
    leftFromText: t('w:leftFromText'),
    rightFromText: t('w:rightFromText'),
    topFromText: t('w:topFromText'),
    bottomFromText: t('w:bottomFromText'),
  };
}

/** Classify a field instruction. */
export function analyzeField(instr: string): { kind?: 'page' | 'numpages' | 'sectionpages'; format?: string; link?: Link } {
  const s = instr.trim();
  const m = /^(\S+)/.exec(s);
  const name = (m?.[1] ?? '').toUpperCase();
  const fmt = /\\\*\s*(\w+)/.exec(s)?.[1];
  const format = fmt && fmt.toUpperCase() !== 'MERGEFORMAT' && fmt.toUpperCase() !== 'CHARFORMAT' ? fmt : undefined;
  if (name === 'PAGE') return { kind: 'page', format };
  if (name === 'NUMPAGES') return { kind: 'numpages', format };
  if (name === 'SECTIONPAGES') return { kind: 'sectionpages', format };
  if (name === 'HYPERLINK') {
    const link: Link = {};
    const url = /^HYPERLINK\s+"([^"]*)"/i.exec(s) ?? /^HYPERLINK\s+(\S+)/i.exec(s);
    if (url && !url[1].startsWith('\\')) link.url = url[1];
    const anchor = /\\l\s+"([^"]*)"/.exec(s);
    if (anchor) link.anchor = anchor[1];
    return { link };
  }
  return {};
}

export { emu, descendant };
