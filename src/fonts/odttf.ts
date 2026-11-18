/**
 * Word embeds fonts "obfuscated" (ECMA-376 Part 1, 17.8.1): the first 32 bytes are XORed with the
 * font key GUID, byte-reversed.
 */
export function deobfuscateFont(data: Uint8Array, fontKey: string): Uint8Array {
  const hex = fontKey.replace(/[{}-]/g, '');
  if (hex.length !== 32) return data;
  const key = new Uint8Array(16);
  for (let i = 0; i < 16; i++) key[i] = parseInt(hex.slice(30 - i * 2, 32 - i * 2), 16);
  const out = data.slice();
  for (let i = 0; i < 32 && i < out.length; i++) out[i] ^= key[i % 16];
  return out;
}
