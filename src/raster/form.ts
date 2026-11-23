// The drawing operators an SVG picture compiles to (see pdf/svg.ts), run onto the canvas: so a
// page image draws exactly the shapes, colours and strokes the PDF's form holds.

import type { Canvas } from './canvas.js';
import { concat, flattenSegs, meanScale, Polylines, type Matrix } from './geometry.js';
import type { PathSeg } from '../layout/types.js';
import { strokeOutline, type Cap, type Join } from './stroke.js';

interface State {
  ctm: Matrix;
  fill: number;
  stroke: number;
  width: number;
  cap: Cap;
  join: Join;
  miterLimit: number;
  dash: number[];
  phase: number;
}

const CAPS: Cap[] = ['butt', 'round', 'square'];
const JOINS: Join[] = ['miter', 'round', 'bevel'];

/** Run form `content` with `m` taking form space to device pixels. */
export function drawForm(canvas: Canvas, content: string, m: Matrix): void {
  let state: State = { ctm: m, fill: 0, stroke: 0, width: 1, cap: 'butt', join: 'miter', miterLimit: 10, dash: [], phase: 0 };
  const stack: State[] = [];
  const operands: number[] = [];
  let array: number[] | null = null;
  let lastArray: number[] = [];
  let segs: PathSeg[] = [];
  const flat = new Polylines();
  const outline = new Polylines();
  const rgb = () => {
    const [r, g, b] = operands.slice(-3).map((v) => Math.round(Math.max(0, Math.min(1, v)) * 255));
    return (r << 16) | (g << 8) | b;
  };
  const paint = (fill: boolean, evenOdd: boolean, stroke: boolean) => {
    flat.reset();
    flattenSegs(segs, state.ctm, flat);
    if (fill) canvas.fill(flat, evenOdd, state.fill, 1);
    if (stroke) {
      const scale = meanScale(state.ctm);
      outline.reset();
      strokeOutline(flat, {
        width: state.width > 0 ? state.width * scale : 1,
        cap: state.cap,
        join: state.join,
        miterLimit: state.miterLimit,
        dash: state.dash.map((d) => d * scale),
        phase: state.phase * scale,
      }, outline);
      canvas.fill(outline, false, state.stroke, 1);
    }
    segs = [];
  };
  const tokens = /\[|\]|[+-]?(?:\d+\.?\d*|\.\d+)(?:e[+-]?\d+)?|[A-Za-z*]+/g;
  for (let match = tokens.exec(content); match; match = tokens.exec(content)) {
    const t = match[0];
    if (t === '[') {
      array = [];
      continue;
    }
    if (t === ']') {
      lastArray = array ?? [];
      array = null;
      continue;
    }
    const c = t.charCodeAt(0);
    if ((c >= 48 && c <= 57) || c === 43 || c === 45 || c === 46) {
      (array ?? operands).push(Number(t));
      continue;
    }
    const o = operands;
    const at = (k: number) => o[o.length - k] ?? 0;
    switch (t) {
      case 'q':
        stack.push({ ...state, dash: [...state.dash] });
        break;
      case 'Q':
        state = stack.pop() ?? state;
        break;
      case 'cm':
        state.ctm = concat([at(6), at(5), at(4), at(3), at(2), at(1)], state.ctm);
        break;
      case 'rg':
        state.fill = rgb();
        break;
      case 'RG':
        state.stroke = rgb();
        break;
      case 'w':
        state.width = at(1);
        break;
      case 'J':
        state.cap = CAPS[at(1)] ?? 'butt';
        break;
      case 'j':
        state.join = JOINS[at(1)] ?? 'miter';
        break;
      case 'M':
        state.miterLimit = Math.max(1, at(1));
        break;
      case 'd':
        state.dash = lastArray;
        state.phase = at(1);
        break;
      case 'm':
        segs.push({ op: 'M', pts: [at(2), at(1)] });
        break;
      case 'l':
        segs.push({ op: 'L', pts: [at(2), at(1)] });
        break;
      case 'c':
        segs.push({ op: 'C', pts: [at(6), at(5), at(4), at(3), at(2), at(1)] });
        break;
      case 'h':
        segs.push({ op: 'Z', pts: [] });
        break;
      case 'f':
        paint(true, false, false);
        break;
      case 'f*':
        paint(true, true, false);
        break;
      case 'B':
        paint(true, false, true);
        break;
      case 'B*':
        paint(true, true, true);
        break;
      case 'S':
        paint(false, false, true);
        break;
      case 'n':
        segs = [];
        break;
    }
    o.length = 0;
  }
}
