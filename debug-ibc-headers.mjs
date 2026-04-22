import ExcelJS from 'exceljs';

const wb = new ExcelJS.Workbook();
await wb.xlsx.readFile('Ejemplo IBC Mes Anterior.xlsx');
const ws = wb.getWorksheet('Sheet1');

// Print row 17 headers with col numbers and letters
function colLetter(n) {
  let r = '';
  while (n > 0) { n--; r = String.fromCharCode(65 + n % 26) + r; n = Math.floor(n / 26); }
  return r;
}

const row17 = ws.getRow(17);
row17.eachCell({ includeEmpty: false }, (cell, colNum) => {
  const v = cell.value;
  const text = typeof v === 'object' && v !== null
    ? (v.richText ? v.richText.map(r => r.text).join('') : JSON.stringify(v))
    : String(v ?? '');
  if (text.trim()) console.log(`${colLetter(colNum)} (${colNum}): ${text.trim()}`);
});
