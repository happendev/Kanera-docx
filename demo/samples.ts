export interface Sample {
  id: string;
  title: string;
  suite: 'classic' | 'issue';
  featured: boolean;
  docx: string;
  reference: { kind: 'pdf' | 'png'; app: string; files: string[] };
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null;
}

function isSample(value: unknown): value is Sample {
  if (!isRecord(value)) return false;
  const { id, title, suite, featured, docx, reference } = value;
  if (!isRecord(reference)) return false;
  const { kind, app, files } = reference;
  return typeof id === 'string' && typeof title === 'string' &&
    (suite === 'classic' || suite === 'issue') && typeof featured === 'boolean' &&
    typeof docx === 'string' && (kind === 'pdf' || kind === 'png') &&
    typeof app === 'string' && Array.isArray(files) && files.every((file: unknown) => typeof file === 'string');
}

export function parseSamples(value: unknown): Sample[] {
  if (!Array.isArray(value) || !value.every(isSample)) throw new Error('Invalid sample list');
  return value;
}
