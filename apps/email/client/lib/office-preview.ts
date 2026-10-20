import JSZip from 'jszip';

const MAX_ROWS = 500;
const MAX_COLUMNS = 50;
const MAX_SHEETS = 20;
const MAX_XML_BYTES = 8 * 1024 * 1024;
const MAX_DEPTH = 256;

const XLSX_TYPES = ['application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'];
const DOCX_TYPES = ['application/vnd.openxmlformats-officedocument.wordprocessingml.document'];
const CSV_TYPES = ['text/csv', 'application/csv'];

export type OfficePreviewKind = 'xlsx' | 'docx' | 'csv';

export const getOfficePreviewKind = (
  mimeType: string,
  filename: string,
): OfficePreviewKind | null => {
  const type = mimeType.toLowerCase().split(';')[0]?.trim() ?? '';
  const extension = filename.toLowerCase().split('.').pop() ?? '';
  if (XLSX_TYPES.includes(type) || extension === 'xlsx') return 'xlsx';
  if (DOCX_TYPES.includes(type) || extension === 'docx') return 'docx';
  if (CSV_TYPES.includes(type) || extension === 'csv') return 'csv';
  return null;
};

const escapeHtml = (value: string) =>
  value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');

const XML_ENTITIES = new Map([
  ['lt', '<'],
  ['gt', '>'],
  ['quot', '"'],
  ['apos', "'"],
  ['amp', '&'],
]);

const decodeXml = (value: string) =>
  value.replace(/&(#x[0-9a-f]+|#\d+|\w+);/gi, (entity, body: string) => {
    if (!body.startsWith('#')) return XML_ENTITIES.get(body) ?? entity;
    const code = /^#x/i.test(body) ? parseInt(body.slice(2), 16) : parseInt(body.slice(1), 10);
    return code <= 0x10ffff ? String.fromCodePoint(code) : '�';
  });

type XmlNode = { name: string; attributes: Map<string, string>; children: (XmlNode | string)[] };

// Attachments come from any sender, so this must stay linear: no pattern can
// scan past the next `<`, and unclosed or stray tags are tolerated, not searched for.
const XML_TOKEN = /<(\/?)([^\s/<>]+)((?:\s[^<>]*?)?)(\/?)>|([^<]+)/g;
const XML_ATTRIBUTE = /([^\s=]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g;

const localName = (name: string) => name.slice(name.indexOf(':') + 1);

const parseXml = (xml: string) => {
  const root: XmlNode = { name: '', attributes: new Map(), children: [] };
  const stack = [root];
  for (const [, closing, rawName, rawAttributes, selfClosing, text] of xml.matchAll(XML_TOKEN)) {
    const parent = stack.at(-1) ?? root;
    if (text !== undefined) {
      parent.children.push(decodeXml(text));
      continue;
    }
    if (!rawName || rawName.startsWith('?') || rawName.startsWith('!')) continue;
    const name = localName(rawName);
    if (closing) {
      const index = stack.findLastIndex((node) => node.name === name);
      if (index > 0) stack.length = index;
      continue;
    }
    const node: XmlNode = {
      name,
      attributes: new Map(
        [...(rawAttributes ?? '').matchAll(XML_ATTRIBUTE)].map(
          ([, key = '', doubleQuoted, singleQuoted]) => [
            localName(key),
            decodeXml(doubleQuoted ?? singleQuoted ?? ''),
          ],
        ),
      ),
      children: [],
    };
    parent.children.push(node);
    if (!selfClosing && stack.length < MAX_DEPTH) stack.push(node);
  }
  return root;
};

const elements = (node: XmlNode) =>
  node.children.filter((child): child is XmlNode => typeof child !== 'string');

/** Depth-first matches of `name`, without looking inside a match. */
const descendants = (node: XmlNode, name: string): XmlNode[] =>
  elements(node).flatMap((child) => (child.name === name ? [child] : descendants(child, name)));

const plainText = (node: XmlNode): string =>
  node.children.map((child) => (typeof child === 'string' ? child : plainText(child))).join('');

const SKIPPED_TEXT_ELEMENTS = new Set(['pPr', 'rPr', 'rPh', 'Fallback']);

const textOf = (node: XmlNode): string =>
  elements(node)
    .map((child) => {
      if (child.name === 't') return plainText(child);
      if (child.name === 'tab') return '\t';
      if (child.name === 'br' || child.name === 'cr') return '\n';
      // Properties hold tab-stop definitions; phonetic guides and the legacy copy
      // of a textbox would repeat visible text.
      if (SKIPPED_TEXT_ELEMENTS.has(child.name)) return '';
      return textOf(child);
    })
    .join('');

type StreamableZipObject = JSZip.JSZipObject & {
  internalStream: (type: 'uint8array') => JSZip.JSZipStreamHelper<Uint8Array>;
};

// A few KB of DEFLATE can inflate to gigabytes, so stop reading at the cap.
const readCapped = (file: StreamableZipObject) =>
  new Promise<{ chunks: Uint8Array[]; truncated: boolean }>((resolve, reject) => {
    const chunks: Uint8Array[] = [];
    let size = 0;
    const stream = file.internalStream('uint8array');
    stream
      .on('data', (chunk) => {
        if (size >= MAX_XML_BYTES) return;
        chunks.push(chunk.subarray(0, MAX_XML_BYTES - size));
        size += chunk.length;
        if (size >= MAX_XML_BYTES) {
          stream.pause();
          resolve({ chunks, truncated: true });
        }
      })
      .on('error', reject)
      .on('end', () => resolve({ chunks, truncated: false }))
      .resume();
  });

const readXml = async (zip: JSZip, path: string) => {
  const file = zip.file(path) as StreamableZipObject | null;
  if (!file) return { root: parseXml(''), truncated: false };
  const { chunks, truncated } = await readCapped(file);
  const decoder = new TextDecoder();
  const xml = chunks.map((chunk) => decoder.decode(chunk, { stream: true })).join('');
  return { root: parseXml(xml + decoder.decode()), truncated };
};

const columnIndex = (reference: string) => {
  let index = 0;
  for (const char of reference.replace(/[^A-Z]/gi, '').toUpperCase()) {
    index = index * 26 + char.charCodeAt(0) - 64;
  }
  return index - 1;
};

const INCOMPLETE_NOTE =
  '<p class="note">This preview is incomplete. Download the file to see everything.</p>';

const renderTable = (rows: string[][]) =>
  `<table>${rows
    .map((row) => `<tr>${row.map((cell) => `<td>${escapeHtml(cell)}</td>`).join('')}</tr>`)
    .join('')}</table>`;

const parseXlsx = async (data: Uint8Array) => {
  const zip = await JSZip.loadAsync(data);
  const workbook = await readXml(zip, 'xl/workbook.xml');
  const relationships = await readXml(zip, 'xl/_rels/workbook.xml.rels');
  const sharedStringsXml = await readXml(zip, 'xl/sharedStrings.xml');
  const sharedStrings = descendants(sharedStringsXml.root, 'si').map(textOf);
  const targets = new Map(
    descendants(relationships.root, 'Relationship').map((relationship) => [
      relationship.attributes.get('Id') ?? '',
      relationship.attributes.get('Target') ?? '',
    ]),
  );
  const sheets = descendants(workbook.root, 'sheet');
  let incomplete =
    workbook.truncated ||
    relationships.truncated ||
    sharedStringsXml.truncated ||
    sheets.length > MAX_SHEETS;

  const rendered: string[] = [];
  for (const sheet of sheets.slice(0, MAX_SHEETS)) {
    const target = targets.get(sheet.attributes.get('id') ?? '') ?? '';
    const path = target.startsWith('/') ? target.slice(1) : `xl/${target.replace(/^\.?\//, '')}`;
    const { root, truncated } = await readXml(zip, path);
    incomplete ||= truncated;
    const rows: string[][] = [];

    for (const row of descendants(root, 'row')) {
      const rowNumber = Number(row.attributes.get('r'));
      while (Number.isInteger(rowNumber) && rows.length < Math.min(rowNumber - 1, MAX_ROWS)) {
        rows.push([]);
      }
      if (rows.length >= MAX_ROWS) {
        incomplete = true;
        break;
      }
      const cells: string[] = [];
      let column = -1;
      for (const cell of elements(row).filter((child) => child.name === 'c')) {
        const reference = cell.attributes.get('r');
        column = reference ? columnIndex(reference) : column + 1;
        if (column < 0) continue;
        if (column >= MAX_COLUMNS) {
          incomplete = true;
          break;
        }
        const type = cell.attributes.get('t');
        const value = descendants(cell, 'v').map(plainText).join('');
        let text = value;
        if (type === 's') text = sharedStrings[Number(value)] ?? '';
        else if (type === 'inlineStr') text = descendants(cell, 'is').map(textOf).join('');
        else if (type === 'b') text = value === '1' ? 'TRUE' : 'FALSE';
        while (cells.length < column) cells.push('');
        cells[column] = text;
      }
      rows.push(cells);
    }

    rendered.push(
      `<h2>${escapeHtml(sheet.attributes.get('name') ?? 'Sheet')}</h2>${renderTable(rows)}`,
    );
  }
  return (
    (rendered.join('') || '<p class="note">This workbook has no sheets.</p>') +
    (incomplete ? INCOMPLETE_NOTE : '')
  );
};

const cellText = (cell: XmlNode) => descendants(cell, 'p').map(textOf).join('\n');

const renderDocxBlocks = (node: XmlNode): string =>
  elements(node)
    .map((child) => {
      if (child.name === 'tbl') {
        return renderTable(
          descendants(child, 'tr').map((row) => descendants(row, 'tc').map(cellText)),
        );
      }
      if (child.name !== 'p') return renderDocxBlocks(child);
      const text = textOf(child);
      if (!text.trim()) return '';
      const style = descendants(child, 'pStyle')[0]?.attributes.get('val') ?? '';
      const tag = /^(heading|title)/i.test(style) ? 'h2' : 'p';
      return `<${tag}>${escapeHtml(text)}</${tag}>`;
    })
    .join('');

const parseDocx = async (data: Uint8Array) => {
  const zip = await JSZip.loadAsync(data);
  const { root, truncated } = await readXml(zip, 'word/document.xml');
  return (
    (renderDocxBlocks(root) || '<p class="note">This document has no text.</p>') +
    (truncated ? INCOMPLETE_NOTE : '')
  );
};

const parseCsv = (text: string) => {
  const rows: string[][] = [];
  let row: string[] = [];
  let cell = '';
  let quoted = false;
  for (let i = 0; i < text.length; i++) {
    const char = text[i];
    if (quoted) {
      if (char === '"' && text[i + 1] === '"') {
        cell += '"';
        i++;
      } else if (char === '"') quoted = false;
      else cell += char;
    } else if (char === '"') quoted = true;
    else if (char === ',') {
      row.push(cell);
      cell = '';
    } else if (char === '\n' || char === '\r') {
      if (char === '\r' && text[i + 1] === '\n') i++;
      row.push(cell);
      rows.push(row);
      row = [];
      cell = '';
    } else cell += char;
    if (rows.length > MAX_ROWS) break;
  }
  if (cell || row.length) rows.push([...row, cell]);
  const incomplete = rows.length > MAX_ROWS || rows.some((r) => r.length > MAX_COLUMNS);
  return (
    renderTable(rows.slice(0, MAX_ROWS).map((r) => r.slice(0, MAX_COLUMNS))) +
    (incomplete ? INCOMPLETE_NOTE : '')
  );
};

const STYLES = `
body{font:14px/1.5 system-ui,sans-serif;margin:16px;color:#111}
h2{font-size:16px;margin:24px 0 8px}
table{border-collapse:collapse;margin-bottom:8px}
td{border:1px solid #d4d4d8;padding:4px 8px;white-space:pre-wrap;vertical-align:top}
tr:first-child td{background:#f4f4f5;font-weight:600}
.note{color:#71717a}
@media(prefers-color-scheme:dark){body{background:#111;color:#eee}td{border-color:#3f3f46}tr:first-child td{background:#27272a}}
`;

/**
 * Builds a self-contained, script-free HTML page for an Office/CSV attachment.
 * Everything from the file is HTML-escaped. Legacy .xls/.doc are not supported.
 * shortcut: spreadsheet dates show their raw serial number and formulas show cached values.
 */
export const renderOfficePreview = async (
  kind: OfficePreviewKind,
  data: Uint8Array,
  title: string,
) => {
  const body =
    kind === 'xlsx'
      ? await parseXlsx(data)
      : kind === 'docx'
        ? await parseDocx(data)
        : parseCsv(new TextDecoder().decode(data));
  return `<!doctype html><html><head><meta charset="utf-8"><meta http-equiv="Content-Security-Policy" content="default-src 'none'; style-src 'unsafe-inline'"><title>${escapeHtml(title)}</title><style>${STYLES}</style></head><body>${body}</body></html>`;
};
