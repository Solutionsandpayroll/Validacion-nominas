import { unzipSync, strFromU8 } from 'fflate';
import { readFileSync } from 'fs';

const bytes = new Uint8Array(readFileSync('public/Ejemplo Validacion de Nomina.xlsx').buffer);
const files = unzipSync(bytes);
const xml = strFromU8(files['xl/worksheets/sheet24.xml']);

// Find rows 2 and 3 in the original template
const row2 = xml.match(/<row[^>]*r="2"[^>]*>[\s\S]*?<\/row>/);
const row3 = xml.match(/<row[^>]*r="3"[^>]*>[\s\S]*?<\/row>/);
console.log('Row 2:', row2 ? row2[0].substring(0, 500) : 'NOT FOUND');
console.log('\nRow 3:', row3 ? row3[0].substring(0, 500) : 'NOT FOUND');
