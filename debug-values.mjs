import ExcelJS from 'exceljs';

const wb = new ExcelJS.Workbook();
await wb.xlsx.readFile('Ejemplo IBC Mes Anterior.xlsx');
const ws = wb.getWorksheet('Sheet1');

// Check tarifa columns on first data row: src 74(BV), 87(CI), 100(CV), 95(CQ), 105(DA), 107(DC)
const tarifaCols = [74, 87, 100, 95, 105, 107, 109, 111];
let found = false;
ws.eachRow({ includeEmpty: false }, (row, rowNum) => {
  if (found) return;
  if (rowNum < 20) return;
  const colBVal = row.getCell(2).value;
  if (colBVal !== null && !isNaN(Number(colBVal))) {
    found = true;
    for (const col of tarifaCols) {
      const cell = row.getCell(col);
      console.log(`Col ${col}: value=${JSON.stringify(cell.value)}, numFmt=${cell.numFmt}, type=${cell.type}`);
    }
    // Also check a date col
    const dateCell = row.getCell(25); // Fecha ING
    console.log(`Fecha ING col 25: value=${JSON.stringify(dateCell.value)}, numFmt=${dateCell.numFmt}`);
  }
});
