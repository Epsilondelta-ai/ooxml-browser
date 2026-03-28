export function asArray<T>(value: T | T[] | undefined | null): T[] {
  if (Array.isArray(value)) {
    return value;
  }

  if (value === undefined || value === null) {
    return [];
  }

  return [value];
}

export async function toUint8Array(input: ArrayBuffer | Uint8Array | Blob): Promise<Uint8Array> {
  if (input instanceof Uint8Array) {
    return input;
  }

  if (input instanceof ArrayBuffer) {
    return new Uint8Array(input);
  }

  return new Uint8Array(await input.arrayBuffer());
}

export function decodeText(input: Uint8Array): string {
  return new TextDecoder('utf-8').decode(input);
}

export function encodeText(input: string): Uint8Array {
  return new TextEncoder().encode(input);
}

export function normalizePartUri(input: string): string {
  const raw = input.replace(/\\/g, '/');
  const parts = raw.split('/');
  const stack: string[] = [];

  for (const part of parts) {
    if (!part || part === '.') {
      continue;
    }

    if (part === '..') {
      stack.pop();
      continue;
    }

    stack.push(part);
  }

  return `/${stack.join('/')}`;
}

export function partExtension(uri: string): string {
  const lastSegment = uri.split('/').pop() ?? '';
  const extension = lastSegment.includes('.') ? lastSegment.split('.').pop() ?? '' : '';
  return extension.toLowerCase();
}

export function isXmlLikePath(uri: string): boolean {
  return uri.endsWith('.xml') || uri.endsWith('.rels');
}

export function resolveTargetUri(sourceUri: string | 'package', target: string, targetMode: 'Internal' | 'External'): string | null {
  if (targetMode === 'External') {
    return null;
  }

  const normalizedTarget = target.replace(/\\/g, '/');
  if (normalizedTarget.startsWith('/')) {
    return normalizePartUri(normalizedTarget);
  }

  const sourceDirectory = sourceUri === 'package'
    ? '/'
    : `${sourceUri.slice(0, sourceUri.lastIndexOf('/') + 1)}`;

  return normalizePartUri(`${sourceDirectory}${normalizedTarget}`);
}
