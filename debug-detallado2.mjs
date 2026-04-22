import { unzipSync, strFromU8 } from 'fflate';
import { readFileSync } from 'fs';
const bytes = new Uint8Array(readFileSync('public/Ejemplo Validacion de Nomina.xlsx').buffer);
const files = unzipSync(bytes);
const wbXml = strFromU8(files['xl/workbook.xml']);
const relsXml = strFromU8(files['xl/_rels/workbook.xml.rels']);
const sheetElem = wbXml.match(/<sheet [^>]*name="Detallado Mes"[^>]*\/>/);
const rId = sheetElem[0].match(/r:id="([^"]+)"/)[1];
const relElem = relsXml.match(new RegExp('<Relationship [^>]*Id="' + rId + '"[^>]*/>'));
const target = relElem[0].match(/Target="([^"]+)"/)[1];
const sheetPath = 'xl/' + target.replace(/^\.\.\//, '');
const sheetXml = strFromU8(files[sheetPath]);

// Print full row 3 and row 4 XML
const rows = sheetXml.match(/<row r="\d+"[^>]*>[\s\S]*?<\/row>/g) || [];
rows.slice(0, 6).forEach(r => {
  const rNum = r.match(/r="(\d+)"/)[1];
  console.log(`\n=== Row ${rNum} ===`);
  console.log(r);
});
