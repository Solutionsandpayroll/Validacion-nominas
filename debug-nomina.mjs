import ExcelJS from 'exceljs';

const wb = new ExcelJS.Workbook();
await wb.xlsx.readFile('Ejemplo Nómina Marzo.xlsx');

console.log('=== SHEETS ===');
wb.eachSheet(ws => console.log(ws.id, ws.name));

const ws = wb.worksheets[0]; // first sheet
console.log('\n=== ROW 1-5 ===');
for (let r = 1; r <= 6; r++) {
  const row = ws.getRow(r);
  const cells = [];
  row.eachCell({ includeEmpty: false }, (cell, col) => {
    const v = cell.value;
    const text = typeof v === 'object' && v !== null
      ? (v.richText ? v.richText.map(x => x.text).join('') : JSON.stringify(v))
      : String(v ?? '');
    cells.push(`[${col}]${text.substring(0,30)}`);
  });
  console.log(`Row ${r}: ${cells.join(' | ')}`);
}

console.log('\n=== ROW 4 headers (all cols) ===');
const header = ws.getRow(4);
header.eachCell({ includeEmpty: false }, (cell, col) => {
  const v = cell.value;
  const text = typeof v === 'object' && v !== null
    ? (v.richText ? v.richText.map(x => x.text).join('') : JSON.stringify(v))
    : String(v ?? '');
  if (text.trim()) console.log(`  Col ${col}: ${text.trim()}`);
});

console.log('\n=== First 3 data rows ===');
let count = 0;
ws.eachRow({ includeEmpty: false }, (row, rowNum) => {
  if (rowNum <= 4 || count >= 3) return;
  count++;
  const cells = [];
  row.eachCell({ includeEmpty: false }, (cell, col) => {
    const v = cell.value;
    const text = typeof v === 'object' && v !== null
      ? (v.richText ? v.richText.map(x => x.text).join('') : JSON.stringify(v))
      : String(v ?? '');
    cells.push(`[${col}]${text.substring(0,20)}`);
  });
  console.log(`Row ${rowNum}: ${cells.join(' | ')}`);
});
