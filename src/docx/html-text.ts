/** Read visible text from an HTML altChunk without loading external resources. */
export function htmlText(source: string): string {
  // Remove comments first so an apparent closing tag inside a comment stays hidden.
  let pos = 0;
  let uncommented = '';
  for (;;) {
    const start = source.indexOf('<!--', pos);
    if (start < 0) break;
    const end = source.indexOf('-->', start + 4);
    if (end < 0) break;
    uncommented += source.slice(pos, start);
    pos = end + 3;
  }
  if (pos) source = uncommented + source.slice(pos);

  const tagName = /<(\/?)([a-z!][a-z0-9_]*)/iy;
  const delimiter = /[>"']/g;
  const rawEnds = new Map([['head', /<\/head\s*>/ig], ['script', /<\/script\s*>/ig], ['style', /<\/style\s*>/ig]]);
  let text = '';
  let space = false;
  const append = (s: string) => {
    if (!s) return;
    const collapsed = s.replace(/\s+/g, ' ');
    text += space && collapsed[0] === ' ' ? collapsed.slice(1) : collapsed;
    space = collapsed.endsWith(' ');
  };
  let i = 0;
  while (i < source.length) {
    const start = source.indexOf('<', i);
    if (start < 0) break;
    append(source.slice(i, start));
    tagName.lastIndex = start;
    const tag = tagName.exec(source);
    if (!tag) {
      text += '<';
      space = false;
      i = start + 1;
      continue;
    }
    delimiter.lastIndex = tagName.lastIndex;
    let next = delimiter.exec(source);
    while (next && next[0] !== '>') {
      const quote = source.indexOf(next[0], delimiter.lastIndex);
      if (quote < 0) break;
      delimiter.lastIndex = quote + 1;
      next = delimiter.exec(source);
    }
    if (!next || next[0] !== '>') {
      i = start;
      break;
    }
    i = delimiter.lastIndex;
    const closing = tag[1] === '/';
    const name = tag[2].toLowerCase();
    const end = !closing && rawEnds.get(name);
    if (end) {
      end.lastIndex = i;
      if (end.test(source)) {
        i = end.lastIndex;
        continue;
      }
      // Later starts of this tag cannot have a closing match either.
      rawEnds.delete(name);
    }
    if (/^(?:p|div|li|h[1-6]|br|tr|table|section|article)$/.test(name)) text += '\n';
    else if (closing && (name === 'td' || name === 'th')) text += ' ';
    space = false;
  }
  append(source.slice(i));
  return text;
}
