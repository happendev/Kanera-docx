/** Excel number formats, as charts label their values and axes. */

const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
const DAYS = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];

function general(v: number): string {
  if (v === 0) return '0';
  const a = Math.abs(v);
  if (a >= 1e11 || a < 1e-9) {
    const [m, e] = v.toExponential(5).split('e');
    const exp = Number(e);
    return `${Number(m)}E${exp < 0 ? '-' : '+'}${String(Math.abs(exp)).padStart(2, '0')}`;
  }
  // Excel shows up to about ten significant digits.
  return String(Number(v.toPrecision(10)));
}

function splitSections(code: string): string[] {
  const out: string[] = [];
  let cur = '';
  let quoted = false;
  for (let i = 0; i < code.length; i++) {
    const c = code[i];
    if (c === '"') quoted = !quoted;
    if (c === '\\' && i + 1 < code.length) {
      cur += c + code[++i];
      continue;
    }
    if (c === ';' && !quoted) {
      out.push(cur);
      cur = '';
    } else cur += c;
  }
  out.push(cur);
  return out;
}

/** Literal text of a format fragment: quoted strings and escapes kept, codes like _) and *  dropped. */
function literal(s: string): string {
  let out = '';
  for (let i = 0; i < s.length; i++) {
    const c = s[i];
    if (c === '"') {
      const j = s.indexOf('"', i + 1);
      out += s.slice(i + 1, j < 0 ? s.length : j);
      i = j < 0 ? s.length : j;
    } else if (c === '\\') out += s[++i] ?? '';
    else if (c === '_' || c === '*') i++;
    else out += c;
  }
  return out;
}

function formatDate(serial: number, code: string): string {
  // Excel's dates run from serial 0 to the end of 9999; it shows hashes for values outside them.
  if (!(serial >= 0 && serial < 2958466)) return '#####';
  // Shown to the whole second, rounded (Excel's 0.0014 days is 12:02:01 AM).
  const ms = Math.round((serial - 25569) * 86400) * 1000;
  const d = new Date(ms);
  const tokens = code.match(/"[^"]*"|\\.|yyyy|yy|mmmmm|mmmm|mmm|mm|m|dddd|ddd|dd|d|hh|h|ss|s|AM\/PM|am\/pm|A\/P|a\/p|[^"\\ymdhsAa]+|./g) ?? [];
  const hour12 = tokens.some((t) => /^(AM\/PM|am\/pm|A\/P|a\/p)$/.test(t));
  let out = '';
  for (let i = 0; i < tokens.length; i++) {
    const t = tokens[i];
    // m and mm are minutes after an hour or before seconds.
    const prev = tokens.slice(0, i).reverse().find((x) => /^[hsdmy]/.test(x));
    const next = tokens.slice(i + 1).find((x) => /^[hsdmy]/.test(x));
    const minutes = (t === 'm' || t === 'mm') && (prev?.startsWith('h') || next?.startsWith('s'));
    const h = d.getUTCHours();
    switch (t) {
      case 'yyyy':
        out += d.getUTCFullYear();
        break;
      case 'yy':
        out += String(d.getUTCFullYear() % 100).padStart(2, '0');
        break;
      case 'mmmmm':
        out += MONTHS[d.getUTCMonth()][0];
        break;
      case 'mmmm':
        out += MONTHS[d.getUTCMonth()];
        break;
      case 'mmm':
        out += MONTHS[d.getUTCMonth()].slice(0, 3);
        break;
      case 'mm':
        out += String(minutes ? d.getUTCMinutes() : d.getUTCMonth() + 1).padStart(2, '0');
        break;
      case 'm':
        out += minutes ? d.getUTCMinutes() : d.getUTCMonth() + 1;
        break;
      case 'dddd':
        out += DAYS[d.getUTCDay()];
        break;
      case 'ddd':
        out += DAYS[d.getUTCDay()].slice(0, 3);
        break;
      case 'dd':
        out += String(d.getUTCDate()).padStart(2, '0');
        break;
      case 'd':
        out += d.getUTCDate();
        break;
      case 'hh':
        out += String(hour12 ? h % 12 || 12 : h).padStart(2, '0');
        break;
      case 'h':
        out += hour12 ? h % 12 || 12 : h;
        break;
      case 'ss':
        out += String(d.getUTCSeconds()).padStart(2, '0');
        break;
      case 's':
        out += d.getUTCSeconds();
        break;
      case 'AM/PM':
      case 'am/pm':
        out += h < 12 ? 'AM' : 'PM';
        break;
      case 'A/P':
      case 'a/p':
        out += h < 12 ? 'A' : 'P';
        break;
      default:
        out += literal(t);
    }
  }
  return out;
}

/** Format a number with an Excel number format (the common codes: decimals, grouping, percent, dates). */
export function formatNumber(v: number, code: string | undefined): string {
  if (!Number.isFinite(v)) return '';
  const sections = splitSections((code ?? 'General').replace(/\[(\$[^\]]*|[A-Za-z]+\d*|<[^\]]*|>[^\]]*|=[^\]]*)\]/g, (m) => (/^\[\$([^-\]]+)/.test(m) ? /^\[\$([^-\]]+)/.exec(m)![1] : '')));
  let fmt = sections[0];
  let neg = false;
  if (v < 0 && sections.length > 1 && sections[1].trim() !== '') {
    fmt = sections[1];
    neg = true;
  } else if (v === 0 && sections.length > 2 && sections[2].trim() !== '') fmt = sections[2];
  const value = neg ? -v : v;
  if (/^\s*general\s*$/i.test(fmt) || fmt.trim() === '') return general(value);
  const bare = fmt.replace(/"[^"]*"|\\./g, '');
  if (/[ymdhs]/i.test(bare.replace(/General/gi, '')) && !/[0#?]/.test(bare)) return formatDate(value, fmt);
  const first = bare.search(/[0#?]/);
  if (first < 0) return literal(fmt);
  // The numeric core runs from the first to the last digit placeholder (with its point and commas).
  let start = -1;
  let end = -1;
  let quoted = false;
  for (let i = 0; i < fmt.length; i++) {
    const c = fmt[i];
    if (c === '"') quoted = !quoted;
    else if (c === '\\') i++;
    else if (!quoted && /[0#?]/.test(c)) {
      if (start < 0) start = i;
      end = i;
    }
  }
  const core = fmt.slice(start, end + 1);
  const prefix = literal(fmt.slice(0, start));
  const suffix = literal(fmt.slice(end + 1));
  let x = value;
  for (const c of prefix + suffix) if (c === '%') x *= 100;
  const exp = /E[+-]/i.test(suffix) || /E[+-]/i.test(core);
  if (exp) return prefix + general(x) + suffix;
  const dot = core.indexOf('.');
  const intPart = dot < 0 ? core : core.slice(0, dot);
  const fracPart = dot < 0 ? '' : core.slice(dot + 1);
  const decimals = (fracPart.match(/[0#?]/g) ?? []).length;
  const minFrac = (fracPart.match(/0/g) ?? []).length;
  const minInt = (intPart.match(/0/g) ?? []).length;
  const group = intPart.includes(',');
  let s = Math.abs(x).toFixed(decimals);
  let [ip, fp = ''] = s.split('.');
  while (fp.length > minFrac && fp.endsWith('0')) fp = fp.slice(0, -1);
  if (ip === '0' && minInt === 0) ip = '';
  ip = ip.padStart(minInt, '0');
  if (group) ip = ip.replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  s = ip + (fp ? '.' + fp : '');
  if (!s) s = '0';
  const sign = x < 0 && !neg && Number(s.replace(/[^\d]/g, '')) !== 0 ? '-' : '';
  return sign + prefix + s + suffix;
}
