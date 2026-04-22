import { unzipSync, strFromU8 } from 'fflate';
import { readFileSync } from 'fs';

const bytes = new Uint8Array(readFileSync('public/Ejemplo Validacion de Nomina.xlsx').buffer);
const files = unzipSync(bytes);
const wbXml = strFromU8(files['xl/workbook.xml']);
const m = wbXml.match(/<sheet [^>]*name="Planilla Mes anterior"[^>]*\/>/);
console.log('Sheet element:', m ? m[0] : 'NOT FOUND');

const relsXml = strFromU8(files['xl/_rels/workbook.xml.rels']);
const rIdMatch = m ? m[0].match(/r:id="([^"]+)"/) : null;
const rId = rIdMatch ? rIdMatch[1] : null;
console.log('rId:', rId);

if (rId) {
  const relMatch = relsXml.match(new RegExp('<Relationship [^>]*Id="' + rId + '"[^>]*/>' ));
  console.log('Rel:', relMatch ? relMatch[0] : 'NOT FOUND');
  const targetMatch = relMatch ? relMatch[0].match(/Target="([^"]+)"/) : null;
  const target = targetMatch ? targetMatch[1] : null;
  const path = 'xl/' + target.replace(/^\.\.\//, '');
  console.log('ZIP path:', path);

  const sheetXml = strFromU8(files[path]);
  console.log('\n--- Sheet XML (first 5000 chars) ---');
  console.log(sheetXml.substring(0, 5000));

  // Also print a header row cell to see the format used
  const headerRowMatch = sheetXml.match(/<row[^>]*r="1"[^>]*>[\s\S]*?<\/row>/);
  console.log('\n--- Header row (row 1) ---');
  console.log(headerRowMatch ? headerRowMatch[0].substring(0, 3000) : 'not found');
}
