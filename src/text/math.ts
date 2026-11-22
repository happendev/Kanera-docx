/**
 * Equations as text: LaTeX for Markdown (`$…$`), and a linear form close to Word's own linear
 * format for plain text (x^2, (a+b)/c, √(x)).
 */
import type { MathNode } from '../docx/model.js';
import { mathLetter } from '../fonts/symbol-map.js';

/** Mathematical alphanumerics (𝑥, 𝐀, 𝟐) as the plain letters they style. */
function plainLetters(text: string): string {
  if (!/[\u{1d400}-\u{1d7ff}]/u.test(text)) return text;
  let out = '';
  for (const ch of text) {
    const cp = ch.codePointAt(0)!;
    const m = cp >= 0x1d400 && cp <= 0x1d7ff ? mathLetter(cp) : undefined;
    out += m ? String.fromCodePoint(m.cp) : ch;
  }
  return out;
}

// ---------------------------------------------------------------------------------------------
// Linear text

/** Wrap an argument in parentheses unless it is a single symbol or already bracketed. */
function group(s: string): string {
  if ([...s].length <= 1 || /^\d+(?:\.\d+)?$/.test(s) || /^\p{L}+$/u.test(s) || (/^\(.*\)$/.test(s) && balanced(s.slice(1, -1)))) return s;
  return `(${s})`;
}

function balanced(s: string): boolean {
  let depth = 0;
  for (const ch of s) {
    if (ch === '(') depth++;
    else if (ch === ')' && --depth < 0) return false;
  }
  return depth === 0;
}

/** An equation in linear text: what a reader would type to write it on one line. */
export function mathToText(n: MathNode, eqArr = false): string {
  switch (n.k) {
    case 'text': {
      const text = plainLetters(n.text);
      return eqArr ? text.replace(/&/g, '') : text;
    }
    case 'row':
      return n.items.map((i) => mathToText(i, eqArr)).join('');
    case 'frac': {
      const num = mathToText(n.num);
      const den = mathToText(n.den);
      if (n.type === 'noBar') return `(${num} ${den})`;
      return `${group(num)}/${group(den)}`;
    }
    case 'script': {
      const base = mathToText(n.base);
      const sub = n.sub ? '_' + group(mathToText(n.sub)) : '';
      const sup = n.sup ? '^' + group(mathToText(n.sup)) : '';
      return n.pre ? `${sub}${sup}${base}` : `${base}${sub}${sup}`;
    }
    case 'rad': {
      const e = group(mathToText(n.e));
      const deg = n.deg ? mathToText(n.deg) : '';
      if (!deg || deg === '2') return '√' + e;
      if (deg === '3') return '∛' + e;
      if (deg === '4') return '∜' + e;
      return `${group(deg)}√${e}`;
    }
    case 'nary': {
      const sub = n.sub ? '_' + group(mathToText(n.sub)) : '';
      const sup = n.sup ? '^' + group(mathToText(n.sup)) : '';
      const e = mathToText(n.e);
      return `${n.chr}${sub}${sup}${e ? ' ' + e : ''}`;
    }
    case 'delim':
      return n.beg + n.items.map((i) => mathToText(i)).join(n.sep) + n.end;
    case 'func': {
      const name = mathToText(n.name);
      const e = mathToText(n.e);
      return /^[([{|]/.test(e) ? name + e : `${name} ${e}`;
    }
    case 'limit': {
      const e = mathToText(n.e);
      const lim = group(mathToText(n.lim));
      return n.over ? `${e}^${lim}` : `${e}_${lim}`;
    }
    case 'accent': {
      const e = mathToText(n.e);
      // A combining accent sits on a single letter; a wider base keeps it after the group.
      return [...e].length === 1 ? e + n.chr : `${group(e)}${n.chr}`;
    }
    case 'bar': {
      const e = mathToText(n.e);
      return [...e].length === 1 ? e + (n.over ? '\u0305' : '\u0332') : e;
    }
    case 'group':
      return mathToText(n.e);
    case 'matrix':
      if (n.eqArr) return n.rows.map((r) => r.map((c) => mathToText(c, true)).join(' ')).join('\n');
      return n.rows.map((r) => r.map((c) => mathToText(c)).join(' ')).join('; ');
    case 'box':
      return mathToText(n.e);
    case 'phantom':
      return n.show ? mathToText(n.e) : '';
  }
}

// ---------------------------------------------------------------------------------------------
// LaTeX

const SYMBOLS: Record<string, string> = {
  α: '\\alpha', β: '\\beta', γ: '\\gamma', δ: '\\delta', ε: '\\varepsilon', ϵ: '\\epsilon', ζ: '\\zeta', η: '\\eta', θ: '\\theta', ϑ: '\\vartheta',
  ι: '\\iota', κ: '\\kappa', λ: '\\lambda', μ: '\\mu', ν: '\\nu', ξ: '\\xi', π: '\\pi', ϖ: '\\varpi', ρ: '\\rho', ϱ: '\\varrho', σ: '\\sigma',
  ς: '\\varsigma', τ: '\\tau', υ: '\\upsilon', φ: '\\varphi', ϕ: '\\phi', χ: '\\chi', ψ: '\\psi', ω: '\\omega',
  Γ: '\\Gamma', Δ: '\\Delta', Θ: '\\Theta', Λ: '\\Lambda', Ξ: '\\Xi', Π: '\\Pi', Σ: '\\Sigma', Υ: '\\Upsilon', Φ: '\\Phi', Ψ: '\\Psi', Ω: '\\Omega',
  '∞': '\\infty', '∂': '\\partial', '∇': '\\nabla', '∅': '\\emptyset', ℏ: '\\hbar', ℓ: '\\ell', ℜ: '\\Re', ℑ: '\\Im', ℵ: '\\aleph',
  '±': '\\pm', '∓': '\\mp', '×': '\\times', '÷': '\\div', '·': '\\cdot', '⋅': '\\cdot', '∗': '\\ast', '∘': '\\circ', '•': '\\bullet',
  '≤': '\\le', '≥': '\\ge', '≠': '\\ne', '≈': '\\approx', '≡': '\\equiv', '≅': '\\cong', '∼': '\\sim', '≃': '\\simeq', '∝': '\\propto',
  '≪': '\\ll', '≫': '\\gg', '≺': '\\prec', '≻': '\\succ', '⊥': '\\perp', '∥': '\\parallel', '∣': '\\mid',
  '∈': '\\in', '∉': '\\notin', '∋': '\\ni', '⊂': '\\subset', '⊃': '\\supset', '⊆': '\\subseteq', '⊇': '\\supseteq', '∪': '\\cup', '∩': '\\cap',
  '∧': '\\wedge', '∨': '\\vee', '¬': '\\neg', '∀': '\\forall', '∃': '\\exists', '∄': '\\nexists', '⊕': '\\oplus', '⊗': '\\otimes', '⊙': '\\odot',
  '→': '\\to', '←': '\\leftarrow', '↔': '\\leftrightarrow', '⇒': '\\Rightarrow', '⇐': '\\Leftarrow', '⇔': '\\Leftrightarrow',
  '↦': '\\mapsto', '↑': '\\uparrow', '↓': '\\downarrow', '⟶': '\\longrightarrow', '⟹': '\\Longrightarrow',
  '…': '\\ldots', '⋯': '\\cdots', '⋮': '\\vdots', '⋱': '\\ddots', '′': "'", '″': "''", '°': '^{\\circ}', '∠': '\\angle', '△': '\\triangle',
  '−': '-', '∖': '\\setminus', '⌊': '\\lfloor', '⌋': '\\rfloor', '⌈': '\\lceil', '⌉': '\\rceil', '⟨': '\\langle', '⟩': '\\rangle',
  '∴': '\\therefore', '∵': '\\because', '√': '\\surd', '\u2061': '', '\u2062': '', '\u2063': '', '\u2064': '', '\u200b': '',
};

const TEXT_ESCAPES: Record<string, string> = {
  '#': '\\#', $: '\\$', '%': '\\%', '&': '\\&', _: '\\_', '{': '\\{', '}': '\\}', '~': '\\textasciitilde{}', '^': '\\textasciicircum{}', '\\': '\\textbackslash{}',
};

const ESCAPES: Record<string, string> = { '#': '\\#', $: '\\$', '%': '\\%', '&': '\\&', _: '\\_', '{': '\\{', '}': '\\}', '~': '\\sim', '^': '\\hat{}', '\\': '\\backslash' };

const NARY: Record<string, string> = {
  '∑': '\\sum', '∏': '\\prod', '∐': '\\coprod', '∫': '\\int', '∬': '\\iint', '∭': '\\iiint', '∮': '\\oint', '∯': '\\oiint', '∰': '\\oiiint',
  '⋃': '\\bigcup', '⋂': '\\bigcap', '⋁': '\\bigvee', '⋀': '\\bigwedge', '⨁': '\\bigoplus', '⨂': '\\bigotimes', '⨀': '\\bigodot', '⨄': '\\biguplus', '⨆': '\\bigsqcup',
};

const ACCENTS: Record<string, string> = {
  '\u0302': '\\hat', '\u0303': '\\tilde', '\u0304': '\\bar', '\u0305': '\\bar', '\u0307': '\\dot', '\u0308': '\\ddot', '\u20db': '\\dddot',
  '\u20d7': '\\vec', '\u0301': '\\acute', '\u0300': '\\grave', '\u0306': '\\breve', '\u030c': '\\check', '\u030a': '\\mathring',
  '\u20d6': '\\overleftarrow', '\u20e1': '\\overleftrightarrow',
};

const DELIMS: Record<string, string> = {
  '': '.', '(': '(', ')': ')', '[': '[', ']': ']', '{': '\\{', '}': '\\}', '|': '|', '‖': '\\|', '⟨': '\\langle', '⟩': '\\rangle',
  '⌊': '\\lfloor', '⌋': '\\rfloor', '⌈': '\\lceil', '⌉': '\\rceil', '〈': '\\langle', '〉': '\\rangle', '/': '/',
};

const FUNCTIONS = new Set([
  'sin', 'cos', 'tan', 'cot', 'sec', 'csc', 'arcsin', 'arccos', 'arctan', 'sinh', 'cosh', 'tanh', 'coth', 'log', 'ln', 'lg', 'exp',
  'lim', 'liminf', 'limsup', 'max', 'min', 'sup', 'inf', 'det', 'dim', 'ker', 'deg', 'gcd', 'arg', 'Pr', 'hom',
]);

function latexText(text: string, eqArr: boolean, plain: boolean): string {
  text = plainLetters(text);
  if (plain) {
    // Ordinary text inside an equation.
    const escaped = [...text].map((ch) => TEXT_ESCAPES[ch] ?? ch).join('');
    return text.trim() ? `\\text{${escaped}}` : text.replace(/ /g, '\\ ');
  }
  let out = '';
  for (const ch of text) {
    if (ch === '&' && eqArr) out += '&';
    else if (ch in ESCAPES) out += ESCAPES[ch];
    else if (ch in SYMBOLS) {
      const s = SYMBOLS[ch];
      // A control word followed by a letter needs a space between them.
      out += s;
      if (/\\[a-zA-Z]+$/.test(s)) out += ' ';
    } else out += ch;
  }
  return out;
}

function brace(s: string): string {
  s = s.trim();
  return [...s].length === 1 ? s : `{${s}}`;
}

/** A function name (sin, lim) as LaTeX, or undefined when it is not plain text. */
function functionName(n: MathNode): string | undefined {
  if (n.k !== 'text') return undefined;
  const name = plainLetters(n.text).trim();
  if (FUNCTIONS.has(name)) return `\\${name}`;
  if (/^[A-Za-z]+$/.test(name)) return name.length === 1 ? name : `\\operatorname{${name}}`;
  return undefined;
}

/** An equation as LaTeX math (without the surrounding $ signs). */
export function mathToLatex(n: MathNode, eqArr = false): string {
  return tidy(latex(n, eqArr));
}

function tidy(s: string): string {
  return s.replace(/\s+/g, ' ').replace(/ ([}^_)\]]|\\right)/g, '$1').trim();
}

function latex(n: MathNode, eqArr = false): string {
  switch (n.k) {
    case 'text': {
      if (n.upright && !n.plain) {
        const f = functionName(n);
        if (f && f.startsWith('\\')) return f + ' ';
        const t = latexText(n.text, eqArr, false);
        return /^[A-Za-z]+$/.test(t) ? `\\mathrm{${t}}` : t;
      }
      return latexText(n.text, eqArr, !!n.plain);
    }
    case 'row':
      return n.items.map((i) => latex(i, eqArr)).join('');
    case 'frac': {
      const num = latex(n.num);
      const den = latex(n.den);
      if (n.type === 'lin' || n.type === 'skw') return `${brace(num)}/${brace(den)}`;
      if (n.type === 'noBar') return `\\genfrac{}{}{0pt}{}{${num.trim()}}{${den.trim()}}`;
      return `\\frac{${num.trim()}}{${den.trim()}}`;
    }
    case 'script': {
      const base = latex(n.base);
      const sub = n.sub ? `_${brace(latex(n.sub))}` : '';
      const sup = n.sup ? `^${brace(latex(n.sup))}` : '';
      if (n.pre) return `{}${sub}${sup}${brace(base)}`;
      // A base that ends in a script or is compound takes braces so the scripts attach to all of it.
      const b = base.trim();
      const simple = [...b].length === 1 || /^\\[a-zA-Z]+$/.test(b) || /^\\[a-zA-Z]+\{[^{}]*\}$/.test(b) || (b.lastIndexOf('\\left') === 0 && /\\right.$/.test(b));
      return `${simple ? b : `{${b}}`}${sub}${sup}`;
    }
    case 'rad': {
      const deg = n.deg ? latex(n.deg).trim() : '';
      return deg ? `\\sqrt[${deg}]{${latex(n.e).trim()}}` : `\\sqrt{${latex(n.e).trim()}}`;
    }
    case 'nary': {
      const op = n.chr === '' ? '' : (NARY[n.chr] ?? n.chr);
      const limits = n.under && INTEGRAL_RE.test(n.chr) ? '\\limits' : '';
      const sub = n.sub ? `_${brace(latex(n.sub))}` : '';
      const sup = n.sup ? `^${brace(latex(n.sup))}` : '';
      return `${op || '{}'}${limits}${sub}${sup} ${latex(n.e)}`;
    }
    case 'delim': {
      const beg = DELIMS[n.beg] ?? n.beg;
      const end = DELIMS[n.end] ?? n.end;
      const sep = n.sep === '|' ? ' \\middle| ' : ` ${latexText(n.sep, false, false)} `;
      const inner = n.items.map((i) => latex(i).trim()).join(n.items.length > 1 ? sep : '');
      // Plain round and square brackets read better without \left and \right around simple content.
      if ((beg === '(' && end === ')') || (beg === '[' && end === ']')) {
        if (!/\\frac|\\sum|\\int|\\prod|\\begin|\\genfrac/.test(inner) && !inner.includes('\\middle')) return `${beg}${inner}${end}`;
      }
      return `\\left${beg}${inner}\\right${end}`;
    }
    case 'func': {
      const name = functionName(n.name) ?? latex(n.name);
      const e = latex(n.e);
      return `${name}${name.startsWith('\\') && /^[A-Za-z0-9]/.test(e) ? ' ' : ''}${e}`;
    }
    case 'limit': {
      const lim = latex(n.lim).trim();
      const op = functionName(n.e);
      if (!n.over && op && FUNCTIONS.has(op.slice(1))) return `${op}_{${lim}} `;
      const e = latex(n.e).trim();
      return n.over ? `\\overset{${lim}}{${e}}` : `\\underset{${lim}}{${e}}`;
    }
    case 'accent':
      return `${ACCENTS[n.chr] ?? '\\hat'}{${latex(n.e).trim()}}`;
    case 'bar':
      return n.over ? `\\overline{${latex(n.e).trim()}}` : `\\underline{${latex(n.e).trim()}}`;
    case 'group': {
      const e = latex(n.e).trim();
      if (n.chr === '⏟' || n.chr === '︸') return `\\underbrace{${e}}`;
      if (n.chr === '⏞' || n.chr === '︷') return `\\overbrace{${e}}`;
      const chr = latexText(n.chr, false, false).trim();
      return n.over ? `\\overset{${chr}}{${e}}` : `\\underset{${chr}}{${e}}`;
    }
    case 'matrix': {
      const rows = n.rows.map((r) => r.map((c) => latex(c, n.eqArr).trim()).join(' & ')).join(' \\\\ ');
      if (n.eqArr) return rows.includes('&') ? `\\begin{aligned}${rows}\\end{aligned}` : `\\begin{gathered}${rows}\\end{gathered}`;
      return `\\begin{matrix}${rows}\\end{matrix}`;
    }
    case 'box':
      return n.border ? `\\boxed{${latex(n.e).trim()}}` : latex(n.e);
    case 'phantom':
      return n.show ? latex(n.e) : `\\phantom{${latex(n.e).trim()}}`;
  }
}

const INTEGRAL_RE = /[∫∬∭∮∯∰]/;
