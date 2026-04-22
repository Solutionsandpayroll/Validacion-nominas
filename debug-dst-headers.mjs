import { unzipSync, strFromU8 } from 'fflate';
import { readFileSync } from 'fs';
const bytes = new Uint8Array(readFileSync('public/Ejemplo Validacion de Nomina.xlsx').buffer);
const files = unzipSync(bytes);

const wbXml = strFromU8(files['xl/workbook.xml']);
const relsXml = strFromU8(files['xl/_rels/workbook.xml.rels']);
const sheetElem = wbXml.match(/<sheet [^>]*name="Planilla Mes anterior"[^>]*\/>/);
const rId = sheetElem[0].match(/r:id="([^"]+)"/)[1];
const relElem = relsXml.match(new RegExp('<Relationship [^>]*Id="' + rId + '"[^>]*/>'));
const target = relElem[0].match(/Target="([^"]+)"/)[1];
const sheetPath = 'xl/' + target.replace(/^\.\.\//, '');

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

const row1Match = sheetXml.match(/<row r="1"[^>]*>([\s\S]*?)<\/row>/);
const cellRe = /<c r="([A-Z]+)1"[^>]*t="s"[^>]*><v>(\d+)<\/v><\/c>/g;
let cm;
const headers = [];
while((cm = cellRe.exec(row1Match[1])) !== null) {
  const col = cm[1];
  let colNum = 0;
  for(const ch of col) colNum = colNum * 26 + ch.charCodeAt(0) - 64;
  headers.push({ col: colNum, letter: col, name: ssEntries[parseInt(cm[2])] });
}
headers.filter(h => h.col > 50).forEach(h => console.log(`${h.letter} (${h.col}): ${h.name}`));
