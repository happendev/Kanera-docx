import { child, children, descendant, elements, textContent, type XmlElement } from '../util/xml.js';
import type { MathNode, RunProps } from './model.js';

export interface MathContext {
  /** Formatting of a math run or control character from its w:rPr (Cambria Math unless it names a font). */
  props(rPr: XmlElement | undefined): RunProps;
}

export const INTEGRALS = new Set([...'∫∬∭∮∯∰∱∲∳⨌⨍⨎⨏⨐⨑⨒⨓⨔⨕⨖⨗⨘⨙⨚⨛⨜']);

function val(pr: XmlElement | undefined, name: string): string | undefined {
  return child(pr, name)?.attrs['m:val'];
}

function flag(pr: XmlElement | undefined, name: string): boolean {
  const el = child(pr, name);
  if (!el) return false;
  const v = el.attrs['m:val'];
  return v === undefined || v === '1' || v === 'on' || v === 'true';
}

/** An OMML element (m:oMath or one of its arguments) as a row of math nodes. */
export function parseOmml(el: XmlElement | undefined, ctx: MathContext): MathNode {
  const items: MathNode[] = [];
  for (const c of elements(el)) {
    const n = parseNode(c, ctx);
    if (n?.k === 'row') items.push(...n.items);
    else if (n) items.push(n);
  }
  return items.length === 1 ? items[0] : { k: 'row', items };
}

/** The formatting a structure gives its own characters (m:ctrlPr). */
function control(pr: XmlElement | undefined, ctx: MathContext): RunProps {
  return ctx.props(descendant(child(pr, 'm:ctrlPr'), 'w:rPr'));
}

function argument(el: XmlElement, name: string, ctx: MathContext): MathNode | undefined {
  const arg = child(el, name);
  if (!arg || !elements(arg).some((c) => !c.name.endsWith('Pr'))) return undefined;
  return parseOmml(arg, ctx);
}

function parseNode(el: XmlElement, ctx: MathContext): MathNode | null {
  const pr = elements(el).find((c) => c.name === el.name + 'Pr');
  const e = () => parseOmml(child(el, 'm:e'), ctx);
  switch (el.name) {
    case 'm:r': {
      const text = children(el, 'm:t').map(textContent).join('');
      if (!text) return null;
      const mrPr = child(el, 'm:rPr');
      const sty = val(mrPr, 'm:sty');
      const props = ctx.props(child(el, 'w:rPr'));
      if (sty === 'b' || sty === 'bi') props.bold = true;
      return { k: 'text', text, props, upright: sty === 'p' || sty === 'b' || undefined, plain: flag(mrPr, 'm:nor') || undefined };
    }
    case 'w:r': {
      const text = children(el, 'w:t').map(textContent).join('');
      return text ? { k: 'text', text, props: ctx.props(child(el, 'w:rPr')), plain: true } : null;
    }
    case 'm:f': {
      const type = val(pr, 'm:type');
      return {
        k: 'frac',
        type: type === 'skw' || type === 'lin' || type === 'noBar' ? type : 'bar',
        num: parseOmml(child(el, 'm:num'), ctx),
        den: parseOmml(child(el, 'm:den'), ctx),
        props: control(pr, ctx),
      };
    }
    case 'm:sSup':
    case 'm:sSub':
    case 'm:sSubSup':
    case 'm:sPre':
      return { k: 'script', base: e(), sub: argument(el, 'm:sub', ctx), sup: argument(el, 'm:sup', ctx), pre: el.name === 'm:sPre' || undefined };
    case 'm:rad':
      return { k: 'rad', deg: flag(pr, 'm:degHide') ? undefined : argument(el, 'm:deg', ctx), e: e(), props: control(pr, ctx) };
    case 'm:nary': {
      // An m:chr without a value draws no operator, only the limits.
      const chrEl = child(pr, 'm:chr');
      const chr = chrEl ? (chrEl.attrs['m:val'] ?? '') : '∫';
      const loc = val(pr, 'm:limLoc');
      return {
        k: 'nary',
        chr,
        // Word's defaults: limits beside integrals, above and below other operators.
        under: loc ? loc === 'undOvr' : !INTEGRALS.has(chr),
        sub: flag(pr, 'm:subHide') ? undefined : argument(el, 'm:sub', ctx),
        sup: flag(pr, 'm:supHide') ? undefined : argument(el, 'm:sup', ctx),
        e: e(),
        props: control(pr, ctx),
      };
    }
    case 'm:d':
      return {
        k: 'delim',
        beg: val(pr, 'm:begChr') ?? '(',
        end: val(pr, 'm:endChr') ?? ')',
        sep: val(pr, 'm:sepChr') ?? '|',
        items: children(el, 'm:e').map((c) => parseOmml(c, ctx)),
        props: control(pr, ctx),
      };
    case 'm:func':
      return { k: 'func', name: upright(parseOmml(child(el, 'm:fName'), ctx)), e: e() };
    case 'm:limLow':
    case 'm:limUpp':
      return { k: 'limit', e: e(), lim: parseOmml(child(el, 'm:lim'), ctx), over: el.name === 'm:limUpp' };
    case 'm:acc':
      return { k: 'accent', chr: val(pr, 'm:chr') ?? '̂', e: e(), props: control(pr, ctx) };
    case 'm:bar':
      return { k: 'bar', over: val(pr, 'm:pos') === 'top', e: e(), props: control(pr, ctx) };
    case 'm:groupChr':
      return { k: 'group', chr: val(pr, 'm:chr') ?? '⏟', over: val(pr, 'm:pos') === 'top', e: e(), props: control(pr, ctx) };
    case 'm:m':
      return { k: 'matrix', rows: children(el, 'm:mr').map((r) => children(r, 'm:e').map((c) => parseOmml(c, ctx))), props: control(pr, ctx) };
    case 'm:eqArr':
      return { k: 'matrix', rows: children(el, 'm:e').map((c) => [parseOmml(c, ctx)]), props: control(pr, ctx), eqArr: true };
    case 'm:box':
    case 'm:borderBox':
      return { k: 'box', e: e(), border: el.name === 'm:borderBox', props: control(pr, ctx) };
    case 'm:phant': {
      const show = val(pr, 'm:show');
      return { k: 'phantom', e: e(), show: show !== '0' && show !== 'off' && show !== 'false', zeroW: flag(pr, 'm:zeroWid'), zeroAsc: flag(pr, 'm:zeroAsc'), zeroDesc: flag(pr, 'm:zeroDesc') };
    }
    // Deleted text is gone from the final document; inserted text and wrappers keep their content.
    case 'w:del':
    case 'w:moveFrom':
      return null;
    case 'w:sdt':
      return parseOmml(child(el, 'w:sdtContent'), ctx);
    default:
      return el.name.endsWith('Pr') || !elements(el).length ? null : parseOmml(el, ctx);
  }
}

/** Function names are set upright. */
function upright(n: MathNode): MathNode {
  if (n.k === 'text') return { ...n, upright: true };
  if (n.k === 'row') return { k: 'row', items: n.items.map(upright) };
  if (n.k === 'script') return { ...n, base: upright(n.base) };
  return n;
}

/** Letters Word sets in math italic: Latin and lower-case Greek. */
export function isMathItalic(ch: string): boolean {
  return /^[A-Za-zıȷα-ωϵϑϕϖ]$/.test(ch);
}

/** Word shows a hyphen in an equation as a minus sign, an apostrophe as a prime and an asterisk as the operator. */
export function mathChar(ch: string): string {
  return ch === '-' ? '−' : ch === "'" ? '′' : ch === '*' ? '∗' : ch;
}

/** Combining accents (an accent's m:chr) drawn by their spacing forms. */
export const ACCENT_FORMS: Record<string, string> = {
  '̀': '`', '́': '´', '̂': 'ˆ', '̃': '˜', '̄': '¯', '̅': '¯',
  '̆': '˘', '̇': '˙', '̈': '¨', '̊': '˚', '̌': 'ˇ',
  '⃖': '←', '⃗': '→', '⃡': '↔',
};

/** The characters an equation draws, with the formatting each is drawn in (for loading fonts). */
export function mathTexts(n: MathNode, out: Array<{ props: RunProps; text: string }> = []): Array<{ props: RunProps; text: string }> {
  const add = (props: RunProps, text: string) => text && out.push({ props, text });
  switch (n.k) {
    case 'text': {
      if (n.plain) add(n.props, n.text);
      else {
        const text = [...n.text].map(mathChar).join('').replace(/&/g, '');
        add({ ...n.props, italic: false }, text);
        if (!n.upright && [...text].some(isMathItalic)) add({ ...n.props, italic: true }, text);
      }
      break;
    }
    case 'row':
      for (const it of n.items) mathTexts(it, out);
      break;
    case 'frac':
      add({ ...n.props, italic: false }, n.type === 'lin' ? '/' : '');
      mathTexts(n.num, out), mathTexts(n.den, out);
      break;
    case 'script':
      mathTexts(n.base, out);
      if (n.sub) mathTexts(n.sub, out);
      if (n.sup) mathTexts(n.sup, out);
      break;
    case 'rad':
      if (n.deg) mathTexts(n.deg, out);
      mathTexts(n.e, out);
      break;
    case 'nary':
      add({ ...n.props, italic: false }, n.chr);
      if (n.sub) mathTexts(n.sub, out);
      if (n.sup) mathTexts(n.sup, out);
      mathTexts(n.e, out);
      break;
    case 'delim':
      add({ ...n.props, italic: false }, n.beg + n.sep + n.end);
      for (const it of n.items) mathTexts(it, out);
      break;
    case 'func':
      mathTexts(n.name, out), mathTexts(n.e, out);
      break;
    case 'limit':
      mathTexts(n.e, out), mathTexts(n.lim, out);
      break;
    case 'accent':
      add({ ...n.props, italic: false }, ACCENT_FORMS[n.chr] ?? n.chr);
      mathTexts(n.e, out);
      break;
    case 'group':
      add({ ...n.props, italic: false }, n.chr);
      mathTexts(n.e, out);
      break;
    case 'matrix':
      for (const r of n.rows) for (const c of r) mathTexts(c, out);
      break;
    case 'bar':
    case 'box':
    case 'phantom':
      mathTexts(n.e, out);
      break;
  }
  return out;
}
