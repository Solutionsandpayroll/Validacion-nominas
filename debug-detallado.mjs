import { unzipSync, strFromU8 } from 'fflate';
import { readFileSync } from 'fs';
const bytes = new Uint8Array(readFileSync('public/Ejemplo Validacion de Nomina.xlsx').buffer);
const files = unzipSync(bytes);

const wbXml = strFromU8(files['xl/workbook.xml']);
const relsXml = strFromU8(files['xl/_rels/workbook.xml.rels']);

// Find "Detallado Mes" sheet
const sheetElem = wbXml.match(/<sheet [^>]*name="Detallado Mes"[^>]*\/>/);
if (!sheetElem) { console.log('Sheet not found! Available:'); console.log(wbXml.match(/<sheet [^>]*>/g)); process.exit(); }
const rId = sheetElem[0].match(/r:id="([^"]+)"/)[1];
const relElem = relsXml.match(new RegExp('<Relationship [^>]*Id="' + rId + '"[^>]*/>'));
const target = relElem[0].match(/Target="([^"]+)"/)[1];
const sheetPath = 'xl/' + target.replace(/^\.\.\//, '');
console.log('Sheet path:', sheetPath);

const sheetXml = strFromU8(files[sheetPath]);
const ssXml = strFromU8(files['xl/sharedStrings.xml']);

const ssEntries = [];
const reSS = /<si>([\s\S]*?)<\/si>/g;
let m;
while((m = reSS.exec(ssXml)) !== null) {
  const t = m[1].match(/^<t(?:\s[^>]*)?>([^<]*)<\/t>$/);
  ssEntries.push(t ? t[1].replace(/&amp;/g,'&').replace(/&lt;/g,'<').replace(/&gt;/g,'>').replace(/&quot;/g,'"') : m[1].replace(/<[^>]+>/g,'').trim());
}

function colLetter(n) {
  let r = '';
  while (n > 0) { n--; r = String.fromCharCode(65 + n % 26) + r; n = Math.floor(n / 26); }
  return r;
}

// Show row 1 headers
const row1Match = sheetXml.match(/<row r="1"[^>]*>([\s\S]*?)<\/row>/);
if (row1Match) {
  console.log('\n=== Row 1 headers ===');
  const cellRe = /<c r="([A-Z]+)1"[^>]*(?:t="s"[^>]*><v>(\d+)<\/v>|><v>([^<]*)<\/v>)[^<]*<\/c>/g;
  let cm;
  while((cm = cellRe.exec(row1Match[1])) !== null) {
    const col = cm[1];
    let colNum = 0; for(const ch of col) colNum = colNum * 26 + ch.charCodeAt(0) - 64;
    const val = cm[2] !== undefined ? ssEntries[parseInt(cm[2])] : cm[3];
    console.log(`  ${col} (${colNum}): ${val}`);
  }
}

// Show rows 1-5 raw XML
console.log('\n=== Rows 1-5 raw XML ===');
const allRows = sheetXml.match(/<row r="[1-5]"[^>]*>[\s\S]*?<\/row>/g) || [];
allRows.slice(0,5).forEach(r => console.log(r.substring(0, 500)));
