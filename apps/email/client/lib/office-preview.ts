import JSZip from 'jszip';

const MAX_ROWS = 500;
const MAX_COLUMNS = 50;

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

const decodeXml = (value: string) =>
  value
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&#x([0-9a-f]+);/gi, (_, hex: string) => String.fromCodePoint(parseInt(hex, 16)))
    .replace(/&#(\d+);/g, (_, code: string) => String.fromCodePoint(parseInt(code, 10)))
    .replace(/&amp;/g, '&');

const collectText = (xml: string, tag: string) =>
  [...xml.matchAll(new RegExp(`<${tag}(?:\\s[^>]*)?>([\\s\\S]*?)</${tag}>`, 'g'))]
    .map((match) => decodeXml(match[1] ?? ''))
    .join('');

const columnIndex = (reference: string) => {
  let index = 0;
  for (const char of reference.replace(/[^A-Z]/gi, '').toUpperCase()) {
    index = index * 26 + char.charCodeAt(0) - 64;
  }
  return index - 1;
};

const renderTable = (rows: string[][], truncated: boolean) => {
  const body = rows
    .map((row) => `<tr>${row.map((cell) => `<td>${escapeHtml(cell)}</td>`).join('')}</tr>`)
    .join('');
  const notice = truncated
    ? `<p class="note">Preview limited to the first ${MAX_ROWS} rows and ${MAX_COLUMNS} columns. Download the file to see everything.</p>`
    : '';
  return `<table>${body}</table>${notice}`;
};

const parseXlsx = async (data: Uint8Array) => {
  const zip = await JSZip.loadAsync(data);
  const read = (path: string) => zip.file(path)?.async('string') ?? Promise.resolve('');

  const sharedStrings = [
    ...(await read('xl/sharedStrings.xml')).matchAll(/<si>([\s\S]*?)<\/si>/g),
  ].map((match) => collectText(match[1] ?? '', 't'));

  const relationships = new Map(
    [...(await read('xl/_rels/workbook.xml.rels')).matchAll(/<Relationship\s[^>]*>/g)].map(
      (match) => [
        /\sId="([^"]*)"/.exec(match[0])?.[1] ?? '',
        /\sTarget="([^"]*)"/.exec(match[0])?.[1] ?? '',
      ],
    ),
  );

  const sheets = [...(await read('xl/workbook.xml')).matchAll(/<sheet\s[^>]*>/g)].map((match) => ({
    name: decodeXml(/\sname="([^"]*)"/.exec(match[0])?.[1] ?? 'Sheet'),
    target: relationships.get(/\sr:id="([^"]*)"/.exec(match[0])?.[1] ?? '') ?? '',
  }));

  const rendered: string[] = [];
  for (const sheet of sheets) {
    const path = sheet.target.startsWith('/')
      ? sheet.target.slice(1)
      : `xl/${sheet.target.replace(/^\.?\//, '')}`;
    const xml = await read(path);
    const rows: string[][] = [];
    let truncated = false;

    for (const rowMatch of xml.matchAll(/<row\b[^>]*?(?:\/>|>([\s\S]*?)<\/row>)/g)) {
      if (rows.length >= MAX_ROWS) {
        truncated = true;
        break;
      }
      const row: string[] = [];
      for (const cellMatch of (rowMatch[1] ?? '').matchAll(
        /<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g,
      )) {
        const attributes = cellMatch[1] ?? '';
        const content = cellMatch[2] ?? '';
        const column = columnIndex(/\sr="([^"]*)"/.exec(` ${attributes}`)?.[1] ?? '');
        if (column < 0) continue;
        if (column >= MAX_COLUMNS) {
          truncated = true;
          continue;
        }
        const type = /\st="([^"]*)"/.exec(` ${attributes}`)?.[1];
        const value = collectText(content, 'v');
        let text = value;
        if (type === 's') text = sharedStrings[Number(value)] ?? '';
        else if (type === 'inlineStr') text = collectText(content, 't');
        else if (type === 'b') text = value === '1' ? 'TRUE' : 'FALSE';
        while (row.length < column) row.push('');
        row[column] = text;
      }
      rows.push(row);
    }

    rendered.push(`<h2>${escapeHtml(sheet.name)}</h2>${renderTable(rows, truncated)}`);
  }
  return rendered.join('') || '<p class="note">This workbook has no sheets.</p>';
};

const parseDocx = async (data: Uint8Array) => {
  const zip = await JSZip.loadAsync(data);
  const xml = (await zip.file('word/document.xml')?.async('string')) ?? '';

  const paragraphText = (paragraph: string) =>
    paragraph
      .replace(/<w:tab\s*\/>/g, '\t')
      .replace(/<w:(?:br|cr)\s*\/>/g, '\n')
      .replace(/<(?!\/?w:t[\s>])[^>]*>/g, '')
      .replace(/<w:t(?:\s[^>]*)?>/g, '')
      .replace(/<\/w:t>/g, '');

  const blocks = [...xml.matchAll(/<w:tbl>[\s\S]*?<\/w:tbl>|<w:p\b[^>]*>[\s\S]*?<\/w:p>/g)].map(
    (match) => {
      const block = match[0];
      if (block.startsWith('<w:tbl>')) {
        const rows = [...block.matchAll(/<w:tr\b[^>]*>([\s\S]*?)<\/w:tr>/g)].map((row) =>
          [...(row[1] ?? '').matchAll(/<w:tc>([\s\S]*?)<\/w:tc>/g)].map((cell) =>
            [...(cell[1] ?? '').matchAll(/<w:p\b[^>]*>[\s\S]*?<\/w:p>/g)]
              .map((paragraph) => decodeXml(paragraphText(paragraph[0])))
              .join('\n'),
          ),
        );
        return renderTable(rows, false);
      }
      const text = decodeXml(paragraphText(block));
      const heading = /<w:pStyle\s+w:val="(?:Heading|Title)/i.test(block);
      return text.trim()
        ? `<${heading ? 'h2' : 'p'}>${escapeHtml(text)}</${heading ? 'h2' : 'p'}>`
        : '';
    },
  );
  return blocks.join('') || '<p class="note">This document has no text.</p>';
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
  const truncated = rows.length > MAX_ROWS || rows.some((r) => r.length > MAX_COLUMNS);
  return renderTable(
    rows.slice(0, MAX_ROWS).map((r) => r.slice(0, MAX_COLUMNS)),
    truncated,
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
