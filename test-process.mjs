import ExcelJS from 'exceljs';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { readFileSync, writeFileSync } from 'fs';

// --- Copy of the functions from App.jsx ---

function colToLetter(col) {
  let result = ''
  while (col > 0) {
    col--
    result = String.fromCharCode(65 + (col % 26)) + result
    col = Math.floor(col / 26)
  }
  return result
}

function escapeXml(str) {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
}

function formatDate(d) {
  const dd = String(d.getDate()).padStart(2, '0')
  const mm = String(d.getMonth() + 1).padStart(2, '0')
  return `${dd}/${mm}/${d.getFullYear()}`
}

function buildCell(colNum, rowNum, value) {
  if (value === null || value === undefined || value === '') return ''
  const ref = `${colToLetter(colNum)}${rowNum}`
  if (value instanceof Date) {
    return `<c r="${ref}" t="inlineStr"><is><t xml:space="preserve">${escapeXml(formatDate(value))}</t></is></c>`
  }
  if (typeof value === 'number') {
    if (!isFinite(value)) return ''
    return `<c r="${ref}"><v>${value}</v></c>`
  }
  if (typeof value === 'boolean') {
    return `<c r="${ref}" t="b"><v>${value ? 1 : 0}</v></c>`
  }
  const str = String(value).trim()
  if (!str) return ''
  return `<c r="${ref}" t="inlineStr"><is><t xml:space="preserve">${escapeXml(str)}</t></is></c>`
}

function getCellValue(cell) {
  const v = cell.value
  if (v === null || v === undefined) return null
  if (v instanceof Date) return v
  if (typeof v === 'object') {
    if (v.richText) return v.richText.map((r) => r.text).join('')
    if ('result' in v) return v.result ?? null
    if (v.error) return null
    if (v.text !== undefined) return v.text
    return null
  }
  return v
}

function rebuildSheetData(sheetXml, newRowsXml) {
  if (/<sheetData\s*\/>/.test(sheetXml)) {
    return sheetXml.replace(/<sheetData\s*\/>/, `<sheetData>${newRowsXml}</sheetData>`)
  }
  return sheetXml.replace(
    /(<sheetData>)([\s\S]*?)(<\/sheetData>)/,
    (_, open, content, close) => {
      const headerMatch = content.match(/<row\s[^>]*r="1"[^>]*>[\s\S]*?<\/row>/)
      const headerRow = headerMatch ? headerMatch[0] : ''
      return `${open}${headerRow}${newRowsXml}${close}`
    }
  )
}

const FIELD_MAP = [
  { src: 2,   dst: 1  },
  { src: 3,   dst: 2  },
  { src: 8,   dst: 3  },
  { src: 14,  dst: 10 },
  { src: 16,  dst: 12 },
  { src: 20,  dst: 13 },
  { src: 22,  dst: 15 },
  { src: 24,  dst: 16 },
  { src: 25,  dst: 17 },
  { src: 28,  dst: 18 },
  { src: 29,  dst: 19 },
  { src: 30,  dst: 20 },
  { src: 33,  dst: 21 },
  { src: 34,  dst: 22 },
  { src: 35,  dst: 23 },
  { src: 36,  dst: 24 },
  { src: 37,  dst: 25 },
  { src: 40,  dst: 26 },
  { src: 41,  dst: 27 },
  { src: 49,  dst: 30 },
  { src: 53,  dst: 33 },
  { src: 57,  dst: 36 },
  { src: 60,  dst: 39 },
  { src: 61,  dst: 40 },
  { src: 64,  dst: 43 },
  { src: 77,  dst: 57 },
  { src: 78,  dst: 56 },
  { src: 79,  dst: 58 },
  { src: 80,  dst: 59 },
  { src: 81,  dst: 60 },
  { src: 83,  dst: 62 },
  { src: 91,  dst: 73 },
  { src: 105, dst: 88 },
  { src: 106, dst: 89 },
  { src: 107, dst: 90 },
  { src: 108, dst: 91 },
  { src: 109, dst: 92 },
  { src: 110, dst: 93 },
  { src: 111, dst: 94 },
  { src: 112, dst: 95 },
]

// --- Run the pipeline ---

const ibcWb = new ExcelJS.Workbook();
await ibcWb.xlsx.readFile('Ejemplo IBC Mes Anterior.xlsx');
const ibcSheet = ibcWb.getWorksheet('Sheet1');

const dataRows = [];
ibcSheet.eachRow({ includeEmpty: false }, (row, rowNum) => {
  if (rowNum < 21) return;
  const val = getCellValue(row.getCell(2));
  if (val !== null && val !== '' && !isNaN(Number(val))) {
    dataRows.push(row);
  }
});
console.log(`Data rows found: ${dataRows.length}`);

// Show raw values for first row to diagnose
console.log('\n--- Raw ExcelJS values for row 21 (FIELD_MAP fields) ---');
const firstRow = dataRows[0];
for (const { src, dst } of FIELD_MAP) {
  const cell = firstRow.getCell(src);
  const raw = cell.value;
  const resolved = getCellValue(cell);
  const built = buildCell(dst, 2, resolved);
  const type = raw === null ? 'null' : typeof raw === 'object' ? JSON.stringify(raw).substring(0, 80) : typeof raw;
  console.log(`  src=${src} dst=${dst}: type=${type} | resolved=${JSON.stringify(resolved)} | cell=${built}`);
}

// Build new XML
const newRowsXml = dataRows.map((srcRow, idx) => {
  const rowNum = idx + 2;
  const cells = FIELD_MAP
    .map(({ src, dst }) => buildCell(dst, rowNum, getCellValue(srcRow.getCell(src))))
    .filter(Boolean)
    .join('');
  return cells ? `<row r="${rowNum}">${cells}</row>` : '';
}).filter(Boolean).join('');

console.log('\n--- newRowsXml (first 1000 chars) ---');
console.log(newRowsXml.substring(0, 1000));

// Load template as zip
const bytes = new Uint8Array(readFileSync('public/Ejemplo Validacion de Nomina.xlsx').buffer);
const files = unzipSync(bytes);
const wbXml = strFromU8(files['xl/workbook.xml']);
const relsXml = strFromU8(files['xl/_rels/workbook.xml.rels']);

const sheetMatch = wbXml.match(/<sheet [^>]*name="Planilla Mes anterior"[^>]*\/>/);
const rId = sheetMatch ? (sheetMatch[0].match(/r:id="([^"]+)"/) || [])[1] : null;
const relMatch = relsXml.match(new RegExp('<Relationship [^>]*Id="' + rId + '"[^>]*/>' ));
const target = relMatch ? (relMatch[0].match(/Target="([^"]+)"/) || [])[1] : null;
const sheetPath = 'xl/' + target.replace(/^\.\.\//, '');

const originalSheetXml = strFromU8(files[sheetPath]);
console.log('\n--- Applying rebuildSheetData ---');
const newSheetXml = rebuildSheetData(originalSheetXml, newRowsXml);

// Verify: show the sheetData section of the new XML
const sdStart = newSheetXml.indexOf('<sheetData>');
const sdEnd = newSheetXml.indexOf('</sheetData>');
const newSheetData = newSheetXml.substring(sdStart, sdEnd + 12);
console.log('\n--- New sheetData (first 2000 chars) ---');
console.log(newSheetData.substring(0, 2000));
console.log('\n--- New sheetData (last 500 chars) ---');
console.log(newSheetData.substring(Math.max(0, newSheetData.length - 500)));

// Write output for manual inspection
files[sheetPath] = strToU8(newSheetXml);
delete files['xl/calcChain.xml'];
const outBytes = zipSync(files);
writeFileSync('test-output.xlsx', Buffer.from(outBytes));
console.log('\nWrote test-output.xlsx');

// Also write the raw sheet XML for inspection
writeFileSync('debug-sheet24.xml', newSheetXml);
console.log('Wrote debug-sheet24.xml');
