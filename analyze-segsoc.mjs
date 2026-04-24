import fs from 'fs';
import { unzipSync, strFromU8 } from 'fflate';

const buf = fs.readFileSync('public/Ejemplo Validacion de Nomina.xlsx');
const files = unzipSync(new Uint8Array(buf));
const wbXml = strFromU8(files['xl/workbook.xml']);
const relsXml = strFromU8(files['xl/_rels/workbook.xml.rels']);

// Resolve Seguridad Social sheet path
const sheetM = wbXml.match(/name="Seguridad Social"[^>]*r:id="([^"]+)"/);
const rId = sheetM[1];
const relM = relsXml.match(new RegExp(`Id="${rId}"[^>]*Target="([^"]+)"`));
const sheetPath = 'xl/' + relM[1].replace('../', '');
console.log('Sheet path:', sheetPath);

const sheetXml = strFromU8(files[sheetPath]);

// Count total rows
const allRows = [...sheetXml.matchAll(/<row r="(\d+)"[^>]*>/g)];
console.log('Total rows:', allRows.length);
console.log('Last row number:', allRows[allRows.length - 1][1]);
console.log('First data row:', allRows[0][1]);

// Show rows 3, 4 and 361 in full
for (const targetRow of [3, 4, 361]) {
  const rowM = sheetXml.match(new RegExp(`<row r="${targetRow}"[^>]*>([\\s\\S]*?)<\\/row>`));
  if (rowM) {
    console.log(`\n=== ROW ${targetRow} ===`);
    console.log(rowM[0]);
  } else {
    console.log(`\nROW ${targetRow}: NOT FOUND`);
  }
}
