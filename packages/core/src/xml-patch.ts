export function replaceAttributeValue(source: string, options: { tagName: string; targetAttr: string; newValue: string; keyAttr?: string; keyValue?: string }): string {
  const openTagPattern = new RegExp(`<${escapeRegExp(options.tagName)}\\b[^>]*>`, 'g');
  let replaced = false;

  const result = source.replace(openTagPattern, (tag) => {
    if (replaced) {
      return tag;
    }

    if (options.keyAttr && options.keyValue !== undefined) {
      const keyMatch = new RegExp(`${escapeRegExp(options.keyAttr)}=(["'])(.*?)\\1`).exec(tag);
      if (!keyMatch || keyMatch[2] !== options.keyValue) {
        return tag;
      }
    }

    replaced = true;
    const attrPattern = new RegExp(`(${escapeRegExp(options.targetAttr)}=)(["'])(.*?)\\2`);
    if (attrPattern.test(tag)) {
      return tag.replace(attrPattern, `$1"${escapeXml(options.newValue)}"`);
    }

    return tag.replace(/\/?>(?=$)/, ` ${options.targetAttr}="${escapeXml(options.newValue)}">`);
  });

  return result;
}

export function replaceInnerTextByAttribute(source: string, options: { containerTag: string; textTag: string; newText: string; keyAttr?: string; keyValue?: string; occurrence?: number }): string {
  const containerPattern = new RegExp(`<${escapeRegExp(options.containerTag)}\\b[^>]*>[\\s\\S]*?</${escapeRegExp(options.containerTag)}>`, 'g');
  let matchedCount = 0;

  return source.replace(containerPattern, (container) => {
    if (options.keyAttr && options.keyValue !== undefined) {
      const keyMatch = new RegExp(`${escapeRegExp(options.keyAttr)}=(["'])(.*?)\\1`).exec(container);
      if (!keyMatch || keyMatch[2] !== options.keyValue) {
        return container;
      }
    }

    matchedCount += 1;
    if (options.occurrence !== undefined && matchedCount - 1 !== options.occurrence) {
      return container;
    }

    const textPattern = new RegExp(`(<${escapeRegExp(options.textTag)}\\b[^>]*>)([\\s\\S]*?)(<\\/${escapeRegExp(options.textTag)}>)`);
    return container.replace(textPattern, `$1${escapeXml(options.newText)}$3`);
  });
}

function escapeXml(value: string): string {
  return value
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&apos;');
}

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}
