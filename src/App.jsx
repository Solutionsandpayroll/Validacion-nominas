import { useState, useRef, useEffect } from 'react'
import ExcelJS from 'exceljs'
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate'
import './App.css'

// ─── Utilidades XML / ZIP ────────────────────────────────────────────────────

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
  const yyyy = d.getUTCFullYear()
  const mm = String(d.getUTCMonth() + 1).padStart(2, '0')
  const dd = String(d.getUTCDate()).padStart(2, '0')
  return `${yyyy}-${mm}-${dd}`
}

/**
 * Construye el XML de una celda usando shared strings (t="s") para texto
 * y número puro para valores numéricos.
 * strIdxMap: Map<string, number> con todos los valores de texto → índice en sharedStrings
 */
function buildCell(colNum, rowNum, value, strIdxMap) {
  if (value === null || value === undefined) return ''
  const ref = `${colToLetter(colNum)}${rowNum}`

  if (value instanceof Date) {
    const str = formatDate(value)
    const idx = strIdxMap.get(str)
    if (idx === undefined) return ''
    return `<c r="${ref}" t="s"><v>${idx}</v></c>`
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
  const idx = strIdxMap.get(str)
  if (idx === undefined) return ''
  return `<c r="${ref}" t="s"><v>${idx}</v></c>`
}

/**
 * Lee el sharedStrings.xml y construye un Map<text, 0-based-index>.
 * Sólo indexa las entradas simples (<si><t>...</t></si>).
 */
function parseSharedStrings(ssXml) {
  const map = new Map()
  const re = /<si>([\s\S]*?)<\/si>/g
  let m
  let idx = 0
  while ((m = re.exec(ssXml)) !== null) {
    const inner = m[1]
    // Entrada simple: <t...>texto</t>
    const simple = inner.match(/^<t(?:\s[^>]*)?>([^<]*)<\/t>$/)
    if (simple) {
      // Decodificar entidades básicas para la clave del mapa
      const decoded = simple[1]
        .replace(/&amp;/g, '&')
        .replace(/&lt;/g, '<')
        .replace(/&gt;/g, '>')
        .replace(/&quot;/g, '"')
      map.set(decoded, idx)
    }
    idx++
  }
  return { map, totalCount: idx }
}

/**
 * Agrega strings nuevos al sharedStrings.xml y devuelve el XML actualizado
 * junto con el Map<text, index> completo para todos los valores necesarios.
 */
function buildSharedStringIndex(ssXml, neededStrings) {
  const { map, totalCount } = parseSharedStrings(ssXml)
  const newEntries = []

  for (const str of neededStrings) {
    if (!map.has(str)) {
      const newIdx = totalCount + newEntries.length
      map.set(str, newIdx)
      newEntries.push(`<si><t xml:space="preserve">${escapeXml(str)}</t></si>`)
    }
  }

  let newSsXml = ssXml
  if (newEntries.length > 0) {
    const newTotal = totalCount + newEntries.length
    newSsXml = ssXml
      .replace('</sst>', newEntries.join('') + '</sst>')
      .replace(/uniqueCount="\d+"/, `uniqueCount="${newTotal}"`)
  }

  return { strIdxMap: map, newSsXml }
}

function getAttr(xmlStr, attr) {
  const m = xmlStr.match(new RegExp(`${attr}="([^"]+)"`))
  return m ? m[1] : null
}

/** Mapeo de Element Type (columna D de All Entitlements) al código de concepto (columna D de Novedades) */
const ELEMENT_TYPE_CODE_MAP = new Map([
  ['CO Salaries integrated',                '001050'],
  ['Salary',                                '001050'],
  ['CO Support and maintenance',            '001050'],
  ['Overtime 125%',                         '001061'],
  ['Overtime 175%',                         '001062'],
  ['Overtime 200%',                         '001063'],
  ['Overtime 250%',                         '001064'],
  ['Overtime 35%',                          '001060'],
  ['Overtime 25%',                          '111502'],
  ['Overtime 75%',                          '111503'],
  ['Overtime 100%',                         '001067'],
  ['Overtime Difference Ajust',             '100030'],
  ['Overtime 100% Overpaid',                '200036'],
  ['Holiday',                               '001092'],
  ['Bravo',                                 '100011'],
  ['SIP Payout',                            '100014'],
  ['Overnight',                             '100015'],
  ['Stock Options',                         '110024'],
  ['Advanced Stock Options',               '200024'],
  ['RSU Dividend',                          '100023'],
  ['AEIP Bonus',                            '100007'],
  ['Transactional Bonus',                   '100012'],
  ['Referral Bonus',                        '100018'],
  ['Lump Sum Merit',                        '100013'],
  ['Retention Bonus',                       '100019'],
  ['Sign-On Bonus',                         '100020'],
  ['EBP \u2013 Sensia',                     '100022'],
  ['LATAM Misc Award',                      '100016'],
  ['CO Grossup',                            '100010'],
  ['Relocation Allowance',                  '100001'],
  ['CO Lease Assistance',                   '100002'],
  ['Car Displacement Allowance',            '100004'],
  ['Retroactive Connectivity Benefit',      '100005'],
  ['Libranza Retroactive',                  '100026'],
  ['BANK FEE RETENTION RETURN',             '100006'],
  ['BANK FEE COMISSION',                    '201301'],
  ['Garnishment Return',                    '100009'],
  ['CO Mayor Discounted Value Ret',         '100008'],
  ['Discount Bravo',                        '200016'],
  ['Gross Up Refund',                       '200017'],
  ['Salary Overpaid',                       '200035'],
  ['CO Salary Advance Discount',            '200020'],
  ['Meal Voucher Deduction',                '200011'],
  ['CO Voluntary Tyba',                     '200025'],
  ['AFC Bancoomeva',                        '200008'],
  ['CO Pension contribution Sub Pay',       '102215'],
  ['CO Pension contribution Soli Ded',      '002215'],
  ['CO Pencion Solid-Subsistencia',         '002218'],
  ['CO Severance Definitive',               '001600'],
  ['CO Severance Premium Legal',            '001500'],
  ['CO Severance Act',                      '001530'],
  ['CO Severance Interest Payment',         '001565'],
  ['CO Tax withheld Fuente IMAN',           '003300'],
  ['CO Contribution AFC Colpatria',         '200000'],
  ['CO Contribution AFC Itau',              '200001'],
  ['CO Contribution AFC Occidente',         '200002'],
  ['CO Contribution AFC Davivienda',        '200003'],
  ['CO Contribution AFC Av Villas',         '200004'],
  ['CO Contribution AFC Bogota',            '200005'],
  ['CO Contributions AFC Bancolombia',      '200006'],
  ['CO Contribution AFC BBVA',              '200007'],
  ['CO Contri AFC Banco Caja Social',       '200009'],
  ['CO Contribution AFC FNA',               '200010'],
  ['Housing Assistance',                    '100000'],
  ['CO Medical Reimburstment',              '100003'],
  ['Connectivity Benefit',                  '101500'],
  ['Food Vouchers Payment',                 '111501'],
  ['Vales alimentacion no gravados',        '111500'],
  ['Transportation Allowance',              '101300'],
  ['Libranza Retention',                    '200012'],
  ['Dental Plan',                           '200018'],
  ['Food Vouchers Deduction',               '200205'],
  ['CO GTO NAL Reimburstment',              '200500'],
  ['Media Medicina Prepagada',              '003200'],
  ['Descuento Anticipo no',                 '200019'],
  ['Retencion Fuente Voluntario',           '203300'],
  ['Bonos de Alimentaci\u00f3n Gravados',   '003321'],
  ['Bonos Alimentaci\u00f3n Gravados excede 41 UVT', '003348'],
  ['CO Voluntary Proteccion',               '200510'],
  ['CO Voluntary Skandia',                  '200511'],
  ['CO Voluntary Colfondos',               '200512'],
  ['CO Voluntary Porvenir',                '200513'],
  ['Skandia Patrimonio',                    '200611'],
  ['Vacation',                              '001130'],
  ['Paternity Leave',                       '001153'],
  ['Sick Leave',                            '001150'],
  ['Licencia de duelo',                     '100032'],
  ['Vacation Indemnity',                    '001145'],
  ['Transportation Assistance',             '101300'],
  ['BANK FEE RETENTION',                    '003410'],
])

/** Resuelve el path ZIP de la hoja dado su nombre */
function resolveSheetZipPath(workbookXml, relsXml, sheetName) {
  const escaped = sheetName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
  const sheetElem = workbookXml.match(new RegExp(`<sheet [^>]*name="${escaped}"[^>]*/>`))
  if (!sheetElem) throw new Error(`Hoja "${sheetName}" no encontrada en workbook.xml`)
  const rId = getAttr(sheetElem[0], 'r:id')
  if (!rId) throw new Error(`r:id no encontrado para la hoja "${sheetName}"`)
  const relElem = relsXml.match(new RegExp(`<Relationship [^>]*Id="${rId}"[^>]*/>`))
  if (!relElem) throw new Error(`Relación ${rId} no encontrada en workbook.xml.rels`)
  const target = getAttr(relElem[0], 'Target')
  if (!target) throw new Error('Target no encontrado en la relación')
  return `xl/${target.replace(/^\.\.\//, '')}`
}

/** Reemplaza el contenido de <sheetData> conservando las primeras N filas de cabecera */
function rebuildSheetData(sheetXml, newRowsXml, headerRowCount = 1) {
  const SD_OPEN = '<sheetData>'
  const SD_CLOSE = '</sheetData>'
  const SELF_CLOSE = /<sheetData\s*\/>/

  if (SELF_CLOSE.test(sheetXml)) {
    return sheetXml.replace(SELF_CLOSE, `${SD_OPEN}${newRowsXml}${SD_CLOSE}`)
  }

  const sdStart = sheetXml.indexOf(SD_OPEN)
  const sdEnd = sheetXml.indexOf(SD_CLOSE)
  if (sdStart === -1 || sdEnd === -1) return sheetXml

  const afterOpen = sdStart + SD_OPEN.length
  const content = sheetXml.substring(afterOpen, sdEnd)

  // Encontrar el fin de las primeras headerRowCount filas
  const CLOSE_ROW = '</row>'
  let endPos = 0
  let rowsCounted = 0
  let searchFrom = 0
  while (rowsCounted < headerRowCount) {
    const found = content.indexOf(CLOSE_ROW, searchFrom)
    if (found === -1) break
    endPos = found + CLOSE_ROW.length
    searchFrom = endPos
    rowsCounted++
  }
  const headerRows = content.substring(0, endPos)

  return (
    sheetXml.substring(0, afterOpen) +
    headerRows +
    newRowsXml +
    sheetXml.substring(sdEnd)
  )
}

// ─── Mapeo de columnas ───────────────────────────────────────────────────────
// Mapeo: columna origen IBC (Sheet1, fila 17, primera col de celda combinada) → columna destino (Planilla Mes anterior)
const FIELD_MAP = [
  { src: 2,   dst: 1  }, // No.
  { src: 3,   dst: 2  }, // Tipo ID
  { src: 8,   dst: 3  }, // No ID
  { src: 14,  dst: 10 }, // Tipo de Cotizante
  { src: 16,  dst: 12 }, // Horas Laboradas
  { src: 20,  dst: 13 }, // Extranjero
  { src: 22,  dst: 15 }, // Fecha Radicación en el Exterior
  { src: 24,  dst: 16 }, // ING
  { src: 25,  dst: 17 }, // Fecha ING
  { src: 28,  dst: 18 }, // RET
  { src: 29,  dst: 19 }, // Fecha RET
  { src: 30,  dst: 20 }, // TDE
  { src: 33,  dst: 21 }, // TAE
  { src: 34,  dst: 22 }, // TDP
  { src: 35,  dst: 23 }, // TAP
  { src: 36,  dst: 24 }, // VSP
  { src: 37,  dst: 25 }, // Fecha Inicio VSP
  { src: 40,  dst: 26 }, // VST
  { src: 41,  dst: 27 }, // SLN
  { src: 49,  dst: 30 }, // IGE
  { src: 53,  dst: 33 }, // LMA
  { src: 57,  dst: 36 }, // VAC-LR
  { src: 60,  dst: 39 }, // AVP
  { src: 61,  dst: 40 }, // VCT
  { src: 64,  dst: 43 }, // IRL
  { src: 77,  dst: 57 }, // Cotización Voluntaria Empleador
  { src: 78,  dst: 56 }, // Cotización Voluntaria Afiliado
  { src: 79,  dst: 58 }, // Fondo Solidaridad Pensional
  { src: 80,  dst: 59 }, // Fondo Subsistencia
  { src: 81,  dst: 60 }, // Valor No retenido
  { src: 83,  dst: 62 }, // AFP Destino
  { src: 91,  dst: 73 }, // EPS Destino
  { src: 105, dst: 88 }, // Tarifa SENA
  { src: 106, dst: 89 }, // Valor Cotización SENA
  { src: 107, dst: 90 }, // Tarifa ICBF
  { src: 108, dst: 91 }, // Valor Cotización ICBF
  { src: 109, dst: 92 }, // Tarifa ESAP
  { src: 110, dst: 93 }, // Valor Cotización ESAP
  { src: 111, dst: 94 }, // Tarifa MEN
  { src: 112, dst: 95 }, // Valor Cotización MEN

  // ── Nuevos mapeos ──────────────────────────────────────────
  { src: 21,  dst: 14 }, // Temp. Ext. → Colombiano Temporalmente en el Exterior

  // Fechas SLN
  { src: 42,  dst: 28 }, // Fecha Inicio sln → Inicio SLN
  { src: 46,  dst: 29 }, // Fecha Fin sln → Fin SLN

  // Fechas IGE
  { src: 50,  dst: 31 }, // Fecha Inicio ige → Inicio IGE
  { src: 52,  dst: 32 }, // Fecha Fin ige → Fin IGE

  // Fechas LMA
  { src: 54,  dst: 34 }, // Fecha Inicio lma → Inicio LMA
  { src: 56,  dst: 35 }, // Fecha Fin lma → Fin LMA

  // Fechas VAC-LR
  { src: 58,  dst: 37 }, // Fecha Inicio vac-lr → Inicio VAC-LR
  { src: 59,  dst: 38 }, // Fecha Fin vac-lr → Fin VAC-LR

  // Fechas VCT
  { src: 62,  dst: 41 }, // Fecha Inicio vct → Inicio VCT
  { src: 63,  dst: 42 }, // Fecha Fin vct → Fin VCT

  // Fechas IRL
  { src: 65,  dst: 44 }, // Fecha Inicio irl → Inicio IRL
  { src: 66,  dst: 45 }, // Fecha Fin irl → Fin IRL

  // Salario
  { src: 68,  dst: 47 }, // Valor (BP) → Salario Mensual($)
  { src: 69,  dst: 48 }, // Integral (BQ) → Salario Integral

  // Pensión
  { src: 71,  dst: 50 }, // Administradora (BS) → Administradora AFP
  { src: 72,  dst: 51 }, // Días (BT)
  { src: 73,  dst: 52 }, // IBC (BU)
  { src: 74,  dst: 53 }, // Tarifa (BV)
  { src: 76,  dst: 54 }, // Valor Cotización (BX)
  { src: 75,  dst: 55 }, // Tarifa Alto Riesgo Pensión (BW) → Indicador Alto Riesgo
  { src: 82,  dst: 61 }, // Total (CD)

  // Salud
  { src: 84,  dst: 63 }, // Administradora (CF)
  { src: 85,  dst: 64 }, // Días (CG)
  { src: 86,  dst: 65 }, // IBC (CH)
  { src: 87,  dst: 66 }, // Tarifa (CI)
  { src: 88,  dst: 67 }, // Valor Cotización (CJ)
  { src: 89,  dst: 68 }, // Valor UPC (CK)

  // ARL
  { src: 97,  dst: 74 }, // Administradora (CS)
  { src: 98,  dst: 75 }, // Días (CT)
  { src: 99,  dst: 76 }, // IBC (CU)
  { src: 100, dst: 77 }, // Tarifa (CV)
  { src: 101, dst: 78 }, // Clase Riesgo (CW) → Clase (BZ)
  { src: 102, dst: 81 }, // Valor Cotización (CX) → Valor Cotización (CC)
  { src: 103, dst: 82 }, // Días (CY) → Días (CD)

  // CCF
  { src: 92,  dst: 83 }, // Administradora (CN) → Administradora CCF
  { src: 94,  dst: 84 }, // IBC (CP) → IBC CCF
  { src: 95,  dst: 85 }, // Tarifa (CQ) → Tarifa CCF
  { src: 96,  dst: 86 }, // Valor Cotización (CR) → Valor Cotización CCF

  // Otros
  { src: 104, dst: 87 }, // IBC (CZ) → IBC Otros Parafiscales
  { src: 113, dst: 96 }, // Exonerado SENA e ICBF (DI) → Exonerado parafiscales y salud
]

// Reformatea strings con porcentaje de "16%" / "12.5%" → "16,00%" / "12,50%"
function reformatPercent(v) {
  if (typeof v !== 'string') return v
  const m = v.trim().match(/^(\d+\.?\d*)%$/)
  if (!m) return v
  return parseFloat(m[1]).toFixed(2).replace('.', ',') + '%'
}

// Transformaciones de valor por columna destino
// Primero aplica la regla específica de la columna; si no hay regla, X→SI genérico
// Divide el campo Nombre en hasta 4 partes: PrimerApellido, SegundoApellido, PrimerNombre, SegundoNombre
// Regla: [0]=PrimerApellido (dst4), [1]=SegundoApellido (dst5), [2]=PrimerNombre (dst6), [3+]=SegundoNombre (dst7)
// Con 2 palabras: solo PrimerApellido y PrimerNombre; con 3: sin SegundoNombre
function splitNombre(nombre) {
  if (!nombre) return {}
  const parts = String(nombre).trim().split(/\s+/).filter(Boolean)
  if (parts.length >= 4) return { 4: parts[0], 5: parts[1], 6: parts[2], 7: parts.slice(3).join(' ') }
  if (parts.length === 3) return { 4: parts[0], 5: parts[1], 6: parts[2] }
  if (parts.length === 2) return { 4: parts[0], 6: parts[1] }
  if (parts.length === 1) return { 4: parts[0] }
  return {}
}

function transformValue(v, dst) {
  if (v === null || v === undefined) return v
  const raw = typeof v === 'string' ? v.trim() : v
  switch (dst) {
    case 16: // ING
    case 18: // RET
      if (raw === 'X' || raw === 'x') return 'Todos los sistemas (ARL, AFP, CCF, EPS)'
      return reformatPercent(v)
    case 30: // IGE
      if (raw === 'X' || raw === 'x') return 'INCAPACIDAD GENERAL'
      return reformatPercent(v)
    case 36: // VAC-LR
      if (raw === 'X' || raw === 'x') return 'VACACIONES'
      if (raw === 'L' || raw === 'l') return 'LICENCIA REMUNERADA'
      return reformatPercent(v)
    default:
      if (raw === 'X' || raw === 'x') return 'SI'
      return reformatPercent(v)
  }
}

const getCellValue = (cell) => {
  const v = cell.value
  if (v === null || v === undefined) return null
  if (v instanceof Date) return v
  if (typeof v === 'object') {
    if (v.richText) return v.richText.map((r) => r.text).join('')
    if ('result' in v) return v.result ?? null   // celda con fórmula → usar resultado cacheado
    if (v.error) return null                      // error de fórmula (#REF!, etc.) → vacío
    if (v.text !== undefined) return v.text       // hipervínculo → usar texto
    return null
  }
  return v
}

// Convierte Date JS a número de serie Excel (días desde 30-dic-1899)
function jsDateToExcelSerial(d) {
  return Math.round((d.getTime() / 86400000) + 25569)
}

// Spec por columna (1-indexed) para "Maestro Personal" — type: 'date'|'num'|'str', style: id del template
const MAESTRO_COL_SPEC = {
  1:  { type: 'num', style: 200 }, // A: No.
  2:  { type: 'num', style: 81  }, // B: CODIGO EMPLEADO
  3:  { type: 'str', style: 81  }, // C: CODIGO ALTERNO
  4:  { type: 'str', style: 82  }, // D: TIPO DOC
  5:  { type: 'str', style: 83  }, // E: DOCUMENTO
  6:  { type: 'str', style: 83  }, // F: NOMBRE
  7:  { type: 'date', style: 205 }, // G: F_NACIMI.
  8:  { type: 'str', style: 83  }, // H: SEXO
  9:  { type: 'str', style: 83  }, // I: ESTADO CIVIL
  10: { type: 'num', style: 81  }, // J: # HIJOS
  11: { type: 'num', style: 81  }, // K: PERSONAS A CARGO
  12: { type: 'str', style: 83  }, // L: CIUDAD RESIDENCIA
  13: { type: 'str', style: 83  }, // M: DIRECCIÓN RESIDENCIA
  14: { type: 'str', style: 81  }, // N: TELEFONO
  15: { type: 'str', style: 81  }, // O: CELULAR
  16: { type: 'str', style: 83  }, // P: E-MAIL
  17: { type: 'str', style: 83  }, // Q: PENSIONADO
  18: { type: 'str', style: 81  }, // R: CUENTA GASTO
  19: { type: 'str', style: 83  }, // S: TIPO LIQUIDACIÓN
  20: { type: 'str', style: 83  }, // T: MODO DE LIQUIDACIÓN
  21: { type: 'str', style: 83  }, // U: TIPO COTIZANTE
  22: { type: 'str', style: 83  }, // V: SUBTIPO COTIZANTE
  23: { type: 'str', style: 83  }, // W: VARIABLE
  24: { type: 'str', style: 83  }, // X: PAGO POR DÍAS
  25: { type: 'str', style: 83  }, // Y: EXTRANJERO
  26: { type: 'str', style: 83  }, // Z: RESIDE EN EL EXTRANJERO
  27: { type: 'date', style: 205 }, // AA: F_INGRESO
  28: { type: 'date', style: 205 }, // AB: F_RETIRO
  29: { type: 'str', style: 83  }, // AC: REGIMEN
  30: { type: 'str', style: 83  }, // AD: TIPO_CONTR.
  31: { type: 'str', style: 83  }, // AE: CARGO
  32: { type: 'num', style: 206 }, // AF: HORAS MES
  33: { type: 'str', style: 83  }, // AG: CLASE DE SALARIO
  34: { type: 'num', style: 207 }, // AH: SUELDO
  35: { type: 'num', style: 208 }, // AI: VLR HORA
  36: { type: 'str', style: 83  }, // AJ: BANCO
  37: { type: 'str', style: 83  }, // AK: TIPO CUENTA
  38: { type: 'str', style: 83  }, // AL: N° CUENTA
  39: { type: 'str', style: 83  }, // AM: COMPAÑÍA
  40: { type: 'str', style: 83  }, // AN: SUCURSAL
  41: { type: 'str', style: 83  }, // AO: CENTRO COSTOS
  42: { type: 'str', style: 83  }, // AP: CLASIFICACIÓN NO 1
  43: { type: 'str', style: 83  }, // AQ: CLASIFICACIÓN NO 2
  44: { type: 'str', style: 83  }, // AR: CLASIFICACIÓN NO 3
  45: { type: 'str', style: 83  }, // AS: CLASIFICACIÓN NO 4
  46: { type: 'str', style: 83  }, // AT: CLASIFICACIÓN NO 5
  47: { type: 'str', style: 83  }, // AU: CLASIFICACIÓN NO 6
  48: { type: 'str', style: 83  }, // AV: CLASIFICACIÓN NO 7
  49: { type: 'num', style: 207 }, // AW: DEDUCIBLE VIVIENDA
  50: { type: 'str', style: 83  }, // AX: DEDUCIBLE DEPENDIENTES
  51: { type: 'num', style: 207 }, // AY: DEDUCIBLE MEDICINA PREPAGADA
  52: { type: 'str', style: 83  }, // AZ: METODO RETENCION
  53: { type: 'num', style: 81  }, // BA: % RET.
  54: { type: 'str', style: 83  }, // BB: EPS
  55: { type: 'str', style: 83  }, // BC: AFP
  56: { type: 'str', style: 83  }, // BD: ARL
  57: { type: 'str', style: 83  }, // BE: CCF
  58: { type: 'str', style: 83  }, // BF: FONDO CESANTIAS
  59: { type: 'str', style: 83  }, // BG: AHORRO PROGRAMADO
  60: { type: 'str', style: 83  }, // BH: APORTE VOLUNTARIO
  61: { type: 'str', style: 83  }, // BI: ESTADO
}

function buildMaestroRow(r, srcRow, seqNum, strIdxMap) {
  let cells = `<c r="A${r}" s="200"><v>${seqNum}</v></c>`
  for (let col = 2; col <= 61; col++) {
    const spec = MAESTRO_COL_SPEC[col] || { type: 'str', style: 83 }
    const colLtr = colToLetter(col)
    const raw = getCellValue(srcRow.getCell(col))
    if (raw == null || raw === '') continue
    if (spec.type === 'date' && raw instanceof Date) {
      const serial = jsDateToExcelSerial(raw)
      cells += `<c r="${colLtr}${r}" s="${spec.style}"><v>${serial}</v></c>`
    } else if (spec.type === 'num') {
      const n = typeof raw === 'number' ? raw : Number(raw)
      if (!isNaN(n)) {
        cells += `<c r="${colLtr}${r}" s="${spec.style}"><v>${n}</v></c>`
      } else {
        const s = String(raw).trim()
        if (s) {
          const si = strIdxMap.get(s)
          if (si !== undefined) cells += `<c r="${colLtr}${r}" s="${spec.style}" t="s"><v>${si}</v></c>`
        }
      }
    } else {
      const s = String(raw).trim()
      if (!s) continue
      const si = strIdxMap.get(s)
      if (si !== undefined) cells += `<c r="${colLtr}${r}" s="${spec.style}" t="s"><v>${si}</v></c>`
    }
  }
  return `<row r="${r}" spans="1:61" s="132" customFormat="1" x14ac:dyDescent="0.3">${cells}</row>`
}

/** Construye una fila XML para la hoja "Detallado Mes" */
function buildDetalladoRow(r, empleado, nombre, concepto, descConcepto, cantidad, devengos, deducciones, strIdxMap) {
  const getIdx = (v) => {
    const s = (v == null ? '' : String(v)).trim()
    if (!s) return null
    const idx = strIdxMap.get(s)
    return idx !== undefined ? idx : null
  }
  const empTrimmed = (empleado == null ? '' : String(empleado)).trim()
  const empIsNum = empTrimmed !== '' && /^\d+$/.test(empTrimmed)
  let c = ''
  c += `<c r="A${r}" s="24" t="str"><f>+CONCATENATE(B${r},D${r})</f></c>`
  if (empIsNum) {
    c += `<c r="B${r}" s="84"><v>${empTrimmed}</v></c>`
  } else {
    const idx = getIdx(empTrimmed)
    if (idx !== null) c += `<c r="B${r}" s="84" t="s"><v>${idx}</v></c>`
  }
  const cIdx = getIdx(nombre);      if (cIdx !== null) c += `<c r="C${r}" s="84" t="s"><v>${cIdx}</v></c>`
  const dIdx = getIdx(concepto);     if (dIdx !== null) c += `<c r="D${r}" s="84" t="s"><v>${dIdx}</v></c>`
  const eIdx = getIdx(descConcepto); if (eIdx !== null) c += `<c r="E${r}" s="84" t="s"><v>${eIdx}</v></c>`
  const cant = Number(cantidad ?? 0);   c += `<c r="F${r}" s="85"><v>${isNaN(cant) ? 0 : cant}</v></c>`
  const dev  = Number(devengos ?? 0);   c += `<c r="G${r}" s="85"><v>${isNaN(dev)  ? 0 : dev}</v></c>`
  const ded  = Number(deducciones ?? 0); c += `<c r="H${r}" s="85"><v>${isNaN(ded)  ? 0 : ded}</v></c>`
  c += `<c r="I${r}" s="85"><f>+G${r}-H${r}</f></c>`
  c += `<c r="J${r}" t="str"><f>+IFERROR(CONCATENATE(B${r},VLOOKUP(D${r},Conceptos!A:F,3,0)),0)</f></c>`
  c += `<c r="K${r}" t="str"><f>+IFERROR(CONCATENATE(B${r},VLOOKUP(D${r},Conceptos!A:F,4,0)),0)</f></c>`
  c += `<c r="L${r}" t="str"><f>+IFERROR(CONCATENATE(B${r},VLOOKUP(D${r},Conceptos!A:F,5,0)),0)</f></c>`
  c += `<c r="M${r}" t="str"><f>+IFERROR(CONCATENATE(B${r},VLOOKUP(D${r},Conceptos!A:F,6,0)),0)</f></c>`
  c += `<c r="N${r}" t="str"><f>+IFERROR(CONCATENATE(B${r},VLOOKUP(D${r},Conceptos!A:H,8,0)),0)</f></c>`
  c += `<c r="O${r}" t="str"><f>+IFERROR(CONCATENATE(B${r},VLOOKUP(D${r},Conceptos!A:I,9,0)),0)</f></c>`
  c += `<c r="P${r}" t="str"><f>+IFERROR(CONCATENATE(B${r},VLOOKUP(D${r},Conceptos!A:J,10,0)),0)</f></c>`
  return `<row r="${r}" spans="1:16" ht="15" customHeight="1" x14ac:dyDescent="0.3">${c}</row>`
}

function buildRevIncapRow(r, empleado, nombre, concepto, descConcepto, cantidad, devengos, deducciones, strIdxMap, kRef = '$C$4', cantidadOverride = null, pct = '70%') {
  const getIdx = (v) => {
    const s = (v == null ? '' : String(v)).trim()
    if (!s) return null
    const idx = strIdxMap.get(s)
    return idx !== undefined ? idx : null
  }
  const empTrimmed = (empleado == null ? '' : String(empleado)).trim()
  const empIsNum = empTrimmed !== '' && /^\d+$/.test(empTrimmed)
  let c = ''
  c += `<c r="A${r}" t="str"><f>+CONCATENATE(B${r},D${r})</f></c>`
  if (empIsNum) {
    c += `<c r="B${r}"><v>${empTrimmed}</v></c>`
  } else {
    const idx = getIdx(empTrimmed)
    if (idx !== null) c += `<c r="B${r}" t="s"><v>${idx}</v></c>`
  }
  const cIdx = getIdx(nombre);       if (cIdx !== null) c += `<c r="C${r}" t="s"><v>${cIdx}</v></c>`
  const dIdx = getIdx(concepto);     if (dIdx !== null) c += `<c r="D${r}" t="s"><v>${dIdx}</v></c>`
  const eIdx = getIdx(descConcepto); if (eIdx !== null) c += `<c r="E${r}" t="s"><v>${eIdx}</v></c>`
  const cant = Number(cantidadOverride ?? cantidad ?? 0); c += `<c r="F${r}"><v>${isNaN(cant) ? 0 : cant}</v></c>`
  const dev  = Number(devengos ?? 0);    c += `<c r="G${r}"><v>${isNaN(dev)  ? 0 : dev}</v></c>`
  const ded  = Number(deducciones ?? 0); c += `<c r="H${r}"><v>${isNaN(ded)  ? 0 : ded}</v></c>`
  c += `<c r="I${r}"><f>+VLOOKUP(B${r},'Maestro Personal'!$B$2:$E$312,4,0)</f></c>`
  c += `<c r="J${r}"><f>+VLOOKUP(I${r},'Planilla Mes anterior'!$CX:$CY,2,0)*100/70</f></c>`
  c += `<c r="K${r}"><f>+ROUND(J${r}/30*${pct}*F${r}*${kRef},0)</f></c>`
  c += `<c r="L${r}"><f>+K${r}-G${r}</f></c>`
  c += `<c r="M${r}" t="str"><f>+VLOOKUP(A${r},'Validacion novedades'!$A$5:$G$3497,7,0)</f></c>`
  c += `<c r="N${r}"><f>+M${r}-G${r}</f></c>`
  c += `<c r="O${r}"><f>+VLOOKUP(B${r},'Maestro Personal'!$B$2:$AH$335,33,0)</f></c>`
  return `<row r="${r}" spans="1:15" ht="15" customHeight="1" x14ac:dyDescent="0.3">${c}</row>`
}

// Reemplaza celdas P-S (cols 16-19) en rows 3-15 (zona pivot) con nuevos valores calculados
function overwritePivotInSheet(sheetXml, pivotCellMap) {
  return sheetXml.replace(/<row r="(\d+)"([^>]*)(?<!\/)>([\s\S]*?)<\/row>/g, (match, rStr, attrs, inner) => {
    const r = parseInt(rStr)
    const inOldPivotRange = r >= 3 && r <= 15
    if (!inOldPivotRange && !pivotCellMap.has(r)) return match
    // Eliminar celdas antiguas de columnas P, Q, R, S (cols de una sola letra)
    let cleaned = inner
      .replace(/<c r="[PQRS]\d+"[^>]*\/>/g, '')
      .replace(/<c r="[PQRS]\d+"[^>]*(?<!\/)>([\s\S]*?)<\/c>/g, '')
    const newCells = pivotCellMap.get(r) ?? ''
    // Actualizar spans para incluir columna S (19) si hay celdas nuevas
    let updatedAttrs = attrs
    if (newCells) {
      updatedAttrs = attrs.replace(/\bspans="(\d+):(\d+)"/, (m, s, e) => `spans="${s}:${Math.max(parseInt(e), 19)}"`  )
    }
    return `<row r="${r}"${updatedAttrs}>${cleaned}${newCells}</row>`
  })
}

// ─── Construye una fila XML para la hoja "Validacion novedades" ─────────────
function buildValNovedadesRow(r, empleado, nombre, concepto, descConcepto, cantidad, devengos, deducciones, strIdxMap) {
  const getIdx = (v) => {
    const s = (v == null ? '' : String(v)).trim()
    if (!s) return null
    const idx = strIdxMap.get(s)
    return idx !== undefined ? idx : null
  }
  const empTrimmed = (empleado == null ? '' : String(empleado)).trim()
  const empIsNum = empTrimmed !== '' && /^\d+$/.test(empTrimmed)
  let c = ''
  // Col A: clave concatenada (fórmula), cols B-H: datos, cols I/J/K/L/N: fórmulas
  c += `<c r="A${r}" s="1" t="str"><f>+CONCATENATE(B${r},D${r})</f></c>`
  if (empIsNum) {
    c += `<c r="B${r}" s="84"><v>${empTrimmed}</v></c>`
  } else {
    const idx = getIdx(empTrimmed)
    if (idx !== null) c += `<c r="B${r}" s="84" t="s"><v>${idx}</v></c>`
  }
  const cIdx = getIdx(nombre);       if (cIdx !== null) c += `<c r="C${r}" s="84" t="s"><v>${cIdx}</v></c>`
  const dIdx = getIdx(concepto);     if (dIdx !== null) c += `<c r="D${r}" s="84" t="s"><v>${dIdx}</v></c>`
  const eIdx = getIdx(descConcepto); if (eIdx !== null) c += `<c r="E${r}" s="84" t="s"><v>${eIdx}</v></c>`
  const cant = Number(cantidad ?? 0);    c += `<c r="F${r}" s="85"><v>${isNaN(cant) ? 0 : cant}</v></c>`
  const dev  = Number(devengos ?? 0);    c += `<c r="G${r}" s="85"><v>${isNaN(dev)  ? 0 : dev}</v></c>`
  const ded  = Number(deducciones ?? 0); c += `<c r="H${r}" s="85"><v>${isNaN(ded)  ? 0 : ded}</v></c>`
  c += `<c r="I${r}" s="85"><f>+G${r}-H${r}</f></c>`
  const concTrimmed = (concepto == null ? '' : String(concepto)).trim()
  if (concTrimmed === '001150') {
    c += `<c r="J${r}" s="88"><f>+F${r}</f></c>`
  } else if (concTrimmed === '100015') {
    c += `<c r="J${r}" s="88"><f>+I${r}</f></c>`
  } else {
    c += `<c r="J${r}" s="88"><f>IF(K${r}&lt;&gt;0,K${r},"")</f></c>`
  }
  c += `<c r="K${r}" s="45"><f>IFERROR(VLOOKUP(A${r},Novedades!$H:$I,2,0),0)</f></c>`
  c += `<c r="L${r}" s="46"><f>IFERROR(J${r}-K${r},0)</f></c>`
  c += `<c r="N${r}" s="22"><f>VLOOKUP(B${r},'Maestro Personal'!$B$2:$AA$377,26,0)</f></c>`
  return `<row r="${r}" spans="1:14" ht="12.9" customHeight="1" x14ac:dyDescent="0.3">${c}</row>`
}

// ─── Actualiza los parámetros configurables en la hoja "Parametros" ─────────
/** Fuerza col C = "Incapacidad" en filas de "Conceptos" donde col A es 001150, 001151 o 001162 */
function patchConceptosSheet(sheetXml, ssArr) {
  const TARGET_CODES = new Set(['001150', '001151', '001162'])
  return sheetXml.replace(/<row r="(\d+)"([^>]*)(?<!\/)(>[\s\S]*?)<\/row>/g,
    (match, rStr, attrs, rest) => {
      const r = parseInt(rStr)
      if (r < 2) return match
      const inner = rest.slice(1)
      // Leer valor de col A — soporta shared string (t="s"), inlineStr y valor directo
      const aCell = inner.match(/<c r="A\d+"([^>]*)>(?:<f>[^<]*<\/f>)?(?:<v>([^<]*)<\/v>|<is><t>([^<]*)<\/t><\/is>)/)
      if (!aCell) return match
      const cellAttrs = aCell[1] ?? ''
      const rawVal = aCell[2] ?? aCell[3] ?? ''
      const isShared = /\bt="s"/.test(cellAttrs)
      const aVal = (isShared && ssArr ? (ssArr[parseInt(rawVal)] ?? '') : rawVal).trim()
      if (!TARGET_CODES.has(aVal)) return match
      // Reemplazar o insertar C con inlineStr "Incapacidad"
      const cCell = `<c r="C${r}" t="inlineStr"><is><t>Incapacidad</t></is></c>`
      let newInner = inner.replace(/<c r="C\d+"[^>]*(?<!\/)(>[\s\S]*?<\/c>|\/>)/g, '')
      const inserted = newInner.replace(/(<c r="D\d+")/, `${cCell}$1`)
      newInner = inserted === newInner ? newInner + cCell : inserted
      return `<row r="${r}"${attrs}>${newInner}</row>`
    })
}

function patchParametrosSheet(sheetXml, params) {
  const patchCellV = (xml, cellRef, newVal) =>
    xml.replace(
      new RegExp(`(<c r="${cellRef}"[^>]*(?<!\\/)>)[\\s\\S]*?(<\\/c>)`),
      (_, open, close) => `${open}<v>${newVal}</v>${close}`
    )
  let xml = sheetXml
  xml = patchCellV(xml, 'D13', Math.round(params.uvt))
  for (let i = 0; i < 6; i++) {
    const r = 23 + i
    xml = patchCellV(xml, `B${r}`, Math.round(params.ret[i].desde))
    xml = patchCellV(xml, `C${r}`, params.ret[i].tarifa)
    xml = patchCellV(xml, `D${r}`, Math.round(params.ret[i].descuento))
    xml = patchCellV(xml, `F${r}`, Math.round(params.sol[i].rango))
    xml = patchCellV(xml, `H${r}`, params.sol[i].tarifa)
  }
  return xml
}

/**
 * Lee los registros de la hoja "rete adic" (filas 2+, col A/B/C).
 * Devuelve array de { empleado, nombre, retAdicional }
 */
function parseReteAdicSheet(sheetXml, ssArr) {
  const rows = []
  const cellRe = /<c r="([A-Z]{1,3})(\d+)"([^>]*)(?<!\/)>([\s\S]*?)<\/c>/g
  const byRow = {}
  let m
  while ((m = cellRe.exec(sheetXml)) !== null) {
    const col = m[1], rowNum = parseInt(m[2]), attrs = m[3], inner = m[4]
    if (rowNum < 2) continue  // omitir fila de encabezados
    if (!byRow[rowNum]) byRow[rowNum] = {}
    const vM = inner.match(/<v>([^<]*)<\/v>/)
    if (!vM) continue
    const isStr = attrs.includes('t="s"')
    const val = isStr ? (ssArr[parseInt(vM[1])] ?? '') : parseFloat(vM[1])
    byRow[rowNum][col] = val
  }
  for (const rn of Object.keys(byRow).sort((a, b) => a - b)) {
    const r = byRow[rn]
    if (r.A === undefined && r.B === undefined) continue
    rows.push({
      empleado:     r.A ?? '',
      nombre:       r.B ?? '',
      retAdicional: r.C ?? 0,
    })
  }
  return rows
}

/**
 * Reconstruye la hoja "rete adic" completa con los registros dados.
 * Preserva la fila 1 (encabezados) y reemplaza todo el sheetData.
 * Estilos tomados del template: s="121" A/B, s="122" C
 */
function buildReteAdicSheet(sheetXml, records, ssForwardMap) {
  // Construir nueva sheetData
  const headerRow = sheetXml.match(/<row r="1"[\s\S]*?<\/row>/)
  const header = headerRow ? headerRow[0] : ''

  const dataRows = records.map((rec, i) => {
    const r = i + 2
    const empVal = typeof rec.empleado === 'number' || /^\d+$/.test(String(rec.empleado).trim())
      ? `<c r="A${r}" s="121"><v>${rec.empleado}</v></c>`
      : (() => {
          const idx = ssForwardMap.get(String(rec.empleado))
          return idx !== undefined ? `<c r="A${r}" s="121" t="s"><v>${idx}</v></c>` : ''
        })()
    const nomIdx = ssForwardMap.get(String(rec.nombre))
    const nomCell = nomIdx !== undefined ? `<c r="B${r}" s="121" t="s"><v>${nomIdx}</v></c>` : ''
    const retVal = Number(rec.retAdicional) || 0
    const retCell = `<c r="C${r}" s="122"><v>${retVal}</v></c>`
    return `<row r="${r}" spans="1:3" x14ac:dyDescent="0.3">${empVal}${nomCell}${retCell}</row>`
  })

  const lastRow = records.length + 1
  const allRows = [header, ...dataRows].join('')
  const newDim = `<dimension ref="A1:C${lastRow}"/>`

  return sheetXml
    .replace(/<dimension ref="[^"]*"\/>/,  newDim)
    .replace(/<sheetData>[\s\S]*?<\/sheetData>/, `<sheetData>${allRows}</sheetData>`)
}

// ─── Función genérica: sobrescribe col A con códigos de empleados ────────────
/**
 * Reemplaza la columna A de cualquier hoja con los códigos de empleados.
 *   · Filas firstRow..lastTplRow: reemplaza/elimina col A existente
 *   · Si hay más empleados que huecos: genera filas extra clonando firstRow
 *     con fórmulas re-numeradas y col A con el código correspondiente
 */
function patchSheetColA(sheetXml, empleadoCodes, firstRow, lastTplRow, aStyle) {
  const total = empleadoCodes.length
  const tplRowStr = String(firstRow)

  // ── Paso 1: parchear filas firstRow..lastTplRow ───────────────────────────
  let result = sheetXml.replace(/<row r="(\d+)"([^>]*)(?<!\/)>([\s\S]*?)<\/row>/g,
    (match, rStr, attrs, inner) => {
      const r = parseInt(rStr)
      if (r < firstRow || r > lastTplRow) return match

      let newInner = inner
        .replace(/<c r="A\d+"([^>]*)(?<!\/)>([\s\S]*?)<\/c>/g, '')
        .replace(/<c r="A\d+"[^>]*\/>/g, '')

      const idx = r - firstRow
      if (idx < total && empleadoCodes[idx]) {
        newInner = `<c r="A${r}" s="${aStyle}"><v>${empleadoCodes[idx]}</v></c>` + newInner
      }
      return `<row r="${r}"${attrs}>${newInner}</row>`
    })

  // ── Paso 2: añadir filas extra si hay más empleados que huecos ───────────
  const TPL_SLOTS = lastTplRow - firstRow + 1
  if (total > TPL_SLOTS) {
    const rowTplRe = new RegExp(`<row r="${tplRowStr}"([^>]*)(?<!\\/)>([\\s\\S]*?)<\\/row>`)
    const rowTplM = sheetXml.match(rowTplRe)
    if (rowTplM) {
      const rowAttrs = rowTplM[1]
      const rowInner = rowTplM[2]
      const newRows = []

      for (let idx = TPL_SLOTS; idx < total; idx++) {
        const r = firstRow + idx
        const code = empleadoCodes[idx]
        if (!code) continue

        let inner = rowInner
          // a) Renumerar refs de celda: r="X{firstRow}" → r="X{r}"
          .replace(new RegExp(`r="([A-Z]{1,3})${tplRowStr}"`, 'g'), (_, col) => `r="${col}${r}"`)
          // b) Quitar atributos de shared-formula, conservar texto
          .replace(/<f ([^>]+)>([\s\S]*?)<\/f>/g, '<f>$2</f>')
          // c) Renumerar referencias de fila relativas en fórmulas
          .replace(new RegExp(`(?<!\\$)([A-Z]{1,3})${tplRowStr}(?!\\d)`, 'g'), (_, col) => `${col}${r}`)
          // d) Eliminar valores cacheados en celdas con fórmula
          .replace(/(<f>[\s\S]*?<\/f>)\s*<v>[^<]*<\/v>/g, '$1')
          // e) Reemplazar col A con el nuevo código
          .replace(/<c r="A[^"]*"([^>]*)(?<!\/)>([\s\S]*?)<\/c>/,
            `<c r="A${r}" s="${aStyle}"><v>${code}</v></c>`)
          .replace(/<c r="A[^"]*"[^>]*\/>/,
            `<c r="A${r}" s="${aStyle}"><v>${code}</v></c>`)

        newRows.push(`<row r="${r}"${rowAttrs}>${inner}</row>`)
      }

      result = result.replace('</sheetData>', newRows.join('') + '</sheetData>')
    }
  }

  return result
}

// ─── Actualiza col A de "Seguridad Social" con códigos de empleados ─────────
/** Wrapper de patchSheetColA para "Seguridad Social" (filas 3-361, s="81") */
function patchSegSocialColA(sheetXml, empleadoCodes) {
  return patchSheetColA(sheetXml, empleadoCodes, 3, 361, '81')
}

/** Inyecta/reemplaza la fórmula de col Z en todas las filas >= 4 de "Devengos" (hasta firstRow + count - 1) */
function patchDevengosColZ(sheetXml, count) {
  const lastRow = 4 + count - 1
  return sheetXml.replace(/<row r="(\d+)"([^>]*)(?<!\/)(>[\s\S]*?)<\/row>/g,
    (match, rStr, attrs, rest) => {
      const r = parseInt(rStr)
      if (r < 4 || r > lastRow) return match
      const formula = `+SUMIF('Validacion novedades'!A:A,CONCATENATE(A${r},$Z$2),'Validacion novedades'!F:F)`
      const zCell = `<c r="Z${r}"><f>${formula}</f></c>`
      const inner = rest.slice(1)
        .replace(/<c r="Z\d+"[^>]*(?<!\/)(>[\s\S]*?<\/c>|\/>)/g, '')
      // Insertar antes de la primera celda de col AA en adelante, o al final
      const inserted = inner.replace(/(<c r="A[A-Z]\d+")/, `${zCell}$1`)
      const finalInner = inserted === inner ? inner + zCell : inserted
      return `<row r="${r}"${attrs}>${finalInner}</row>`
    })
}

/** Inyecta/reemplaza la fórmula de col G en todas las filas >= 3 de "Seguridad Social" (hasta firstRow + count - 1) */
function patchSegSocialColG(sheetXml, count) {
  const lastRow = 3 + count - 1
  return sheetXml.replace(/<row r="(\d+)"([^>]*)(?<!\/)(>[\s\S]*?)<\/row>/g,
    (match, rStr, attrs, rest) => {
      const r = parseInt(rStr)
      if (r < 3 || r > lastRow) return match
      const formula = `+IF(Parametros!$D$5="HORAS",TRUNC(SUMIF('Detallado Mes'!J:J,CONCATENATE('Seguridad Social'!A${r},"Incapacidad"),'Detallado Mes'!F:F)/7.3333,0),SUMIF('Detallado Mes'!J:J,CONCATENATE('Seguridad Social'!A${r},"Incapacidad"),'Detallado Mes'!F:F))`
      const gCell = `<c r="G${r}"><f>${formula}</f></c>`
      // Quitar celda G existente (con contenido o self-closing)
      const inner = rest.slice(1, rest.length) // quita el '>' inicial
        .replace(/<c r="G\d+"[^>]*(?<!\/)(>[\s\S]*?<\/c>|\/>)/g, '')
      // Insertar gCell antes de la primera celda de col H en adelante, o al final
      const insertedInner = inner.replace(/(<c r="[H-Z]\d+")/, `${gCell}$1`)
      const finalInner = insertedInner === inner ? inner + gCell : insertedInner
      return `<row r="${r}"${attrs}>${finalInner}</row>`
    })
}

// ─── Extrae hoja "Seguridad Social" como workbook independiente ──────────────
// ─── Forzar recálculo automático en el libro completo ────────────────────────
function addFullCalcOnLoad(wbXml) {
  if (wbXml.includes('<calcPr')) {
    return wbXml.replace(/<calcPr\b([^>]*?)\/>/,
      (_, attrs) => {
        let a = attrs
        a = a.includes('calcMode=') ? a.replace(/calcMode="[^"]*"/, 'calcMode="auto"') : ` calcMode="auto"${a}`
        a = a.includes('fullCalcOnLoad=') ? a.replace(/fullCalcOnLoad="[^"]*"/, 'fullCalcOnLoad="1"') : `${a} fullCalcOnLoad="1"`
        return `<calcPr${a}/>`
      })
  }
  return wbXml.replace('</workbook>', '<calcPr calcMode="auto" fullCalcOnLoad="1"/></workbook>')
}

// ─── HyperFormula: evalúa Seguridad Social y produce XLSX standalone ──────────

/** Construye array inverso de shared strings: índice → texto */
function buildSsReverseArray(ssXml) {
  const arr = []
  const re = /<si>([\s\S]*?)<\/si>/g
  let m
  while ((m = re.exec(ssXml)) !== null) {
    const texts = [...m[1].matchAll(/<t(?:[^>]*)?>([^<]*)<\/t>/g)].map(t =>
      t[1].replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"')
    )
    arr.push(texts.join(''))
  }
  return arr
}

/** Letras de columna a índice base-0: A→0, Z→25, AA→26 */
function colLetterToIdx(s) {
  let n = 0
  for (const ch of s) n = n * 26 + (ch.charCodeAt(0) - 64)
  return n - 1
}

/** Parsea referencia de celda "AB123" → {col, row} (base-0) */
function parseCellRef(ref) {
  const m = ref.match(/^([A-Z]{1,3})(\d+)$/)
  if (!m) return null
  return { col: colLetterToIdx(m[1]), row: parseInt(m[2]) - 1 }
}

/** Ajusta referencias relativas de una fórmula al expandir una shared formula */
function adjustFormula(formula, deltaRow, deltaCol) {
  return formula.replace(/(\$?)([A-Z]{1,3})(\$?)(\d+)/g, (match, colAbs, colLetters, rowAbs, rowNum) => {
    if (colAbs && rowAbs) return match
    let newCol = colLetterToIdx(colLetters)
    let newRow = parseInt(rowNum)
    if (!colAbs) newCol += deltaCol
    if (!rowAbs) newRow += deltaRow
    if (newCol < 0 || newRow < 1) return match
    return colAbs + colToLetter(newCol + 1) + rowAbs + newRow
  })
}

/** Convierte fórmula Excel al formato de HyperFormula (prefijo =) */
function toHFFormula(f) {
  // Desescapar entidades XML (&gt; → >, &lt; → <, &amp; → &, &quot; → ")
  const unescaped = f
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    // Quitar prefijos de libros externos [N] p.ej. [6]Pensionado! → Pensionado!
    .replace(/\[\d+\]/g, '')
  const t = unescaped.trim()
  if (t.startsWith('=')) return t
  if (t.startsWith('+')) return '=' + t.slice(1)
  return '=' + t
}

/**
 * Parsea hoja XML → array 2D de VALORES (usa <v> cacheados, ignora fórmulas).
 * ssArr: string[] — índice → texto (shared strings inverso)
 */
function parseSheetXmlToValues(sheetXml, ssArr) {
  const grid = new Map()
  let maxRow = 0, maxCol = 0
  const rowRe = /<row\b[^>]*r="(\d+)"[^>]*>([\s\S]*?)<\/row>/g
  let rowMatch
  while ((rowMatch = rowRe.exec(sheetXml)) !== null) {
    const r = parseInt(rowMatch[1]) - 1
    if (r > maxRow) maxRow = r
    // (?<!\/) negative lookbehind — excluye celdas self-closing <c r="..."/>
    // Sin esto, <c r="N3" s="39"/> consumiría el contenido de <c r="O3">
    const cellRe = /<c r="([A-Z]{1,3}\d+)"([^>]*)(?<!\/)>([\s\S]*?)<\/c>/g
    let cellMatch
    while ((cellMatch = cellRe.exec(rowMatch[2])) !== null) {
      const pos = parseCellRef(cellMatch[1])
      if (!pos) continue
      if (pos.col > maxCol) maxCol = pos.col
      const t = (cellMatch[2].match(/\bt="([^"]+)"/) || [])[1] || ''
      const inner = cellMatch[3]
      // Soporte para inlineStr
      const isM = t === 'inlineStr' ? inner.match(/<t[^>]*>([^<]*)<\/t>/) : null
      const vM = inner.match(/<v>([^<]*)<\/v>/)
      if (!vM && !isM) continue
      const raw = vM ? vM[1] : ''
      let value
      if (t === 's') value = ssArr[parseInt(raw)] ?? ''
      else if (t === 'b') value = raw === '1'
      else if (t === 'inlineStr') value = (isM ? isM[1] : raw).replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"')
      else if (t === 'str') value = raw.replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"')
      else { const n = parseFloat(raw); value = isNaN(n) ? raw : n }
      grid.set(`${pos.row},${pos.col}`, value)
    }
  }
  const result = []
  for (let r = 0; r <= maxRow; r++) {
    const row = []
    for (let c = 0; c <= maxCol; c++) row.push(grid.get(`${r},${c}`) ?? null)
    result.push(row)
  }
  return result
}

/**
 * Parsea hoja XML → array 2D con FÓRMULAS expandidas (incluyendo shared formulas).
 * Las celdas con fórmula devuelven string con '=' para que HyperFormula las evalúe.
 * Las celdas sin fórmula devuelven el valor cacheado.
 */
function parseSheetXmlToFormulas(sheetXml, ssArr) {
  const grid = new Map()
  const sharedFormulas = new Map() // si → { formula, base: {row, col} }
  let maxRow = 0, maxCol = 0
  const rowRe = /<row\b[^>]*r="(\d+)"[^>]*>([\s\S]*?)<\/row>/g
  let rowMatch
  while ((rowMatch = rowRe.exec(sheetXml)) !== null) {
    const r = parseInt(rowMatch[1]) - 1
    if (r > maxRow) maxRow = r
    // (?<!\/) negative lookbehind — excluye celdas self-closing <c r="..."/>
    const cellRe = /<c r="([A-Z]{1,3}\d+)"([^>]*)(?<!\/)>([\s\S]*?)<\/c>/g
    let cellMatch
    while ((cellMatch = cellRe.exec(rowMatch[2])) !== null) {
      const cellRef = cellMatch[1]
      const pos = parseCellRef(cellRef)
      if (!pos) continue
      if (pos.col > maxCol) maxCol = pos.col
      const attrs = cellMatch[2]
      const inner = cellMatch[3]
      const t = (attrs.match(/\bt="([^"]+)"/) || [])[1] || ''

      // Detectar elemento <f> (con contenido o self-closing)
      const fM = inner.match(/<f([^>\/]*)(?:\/>|>([\s\S]*?)<\/f>)/)
      if (fM) {
        const fAttrs = fM[1]
        const fContent = (fM[2] || '').trim()
        const fType = (fAttrs.match(/\bt="([^"]+)"/) || [])[1] || ''
        const siM = fAttrs.match(/\bsi="(\d+)"/)
        const si = siM ? parseInt(siM[1]) : null
        let formula = null
        if (fContent) {
          if (fType === 'shared' && si !== null) {
            // La base de expansión es el inicio del rango ref=, no necesariamente la celda definidora
            // Ej: <c r="N4" ... t="shared" ref="O4:O66" si="6"> → base debe ser O4
            const refAttrM = fAttrs.match(/\bref="([A-Z]{1,3}\d+)/)
            const refBase = refAttrM ? (parseCellRef(refAttrM[1]) ?? pos) : pos
            sharedFormulas.set(si, { formula: fContent, base: refBase })
          }
          formula = toHFFormula(fContent)
        } else if (fType === 'shared' && si !== null) {
          const master = sharedFormulas.get(si)
          if (master) {
            formula = toHFFormula(adjustFormula(master.formula, pos.row - master.base.row, pos.col - master.base.col))
          }
        }
        if (formula) {
          grid.set(`${pos.row},${pos.col}`, formula)
          continue
        }
      }

      // Valor cacheado como fallback
      const isM = t === 'inlineStr' ? inner.match(/<t[^>]*>([^<]*)<\/t>/) : null
      const vM = inner.match(/<v>([^<]*)<\/v>/)
      if (!vM && !isM) continue
      const raw = vM ? vM[1] : ''
      let value
      if (t === 's') value = ssArr[parseInt(raw)] ?? ''
      else if (t === 'b') value = raw === '1'
      else if (t === 'inlineStr') value = (isM ? isM[1] : raw).replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"')
      else if (t === 'str') value = raw.replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"')
      else { const n = parseFloat(raw); value = isNaN(n) ? raw : n }
      grid.set(`${pos.row},${pos.col}`, value)
    }
  }
  const result = []
  for (let r = 0; r <= maxRow; r++) {
    const row = []
    for (let c = 0; c <= maxCol; c++) row.push(grid.get(`${r},${c}`) ?? null)
    result.push(row)
  }
  return result
}

/**
 * Reemplaza las celdas con fórmula en el XML de sheet19 por sus valores evaluados.
 * evalMap: Map<cellRef, valor>   (ej. "B3" → 12345)
 * ssMap:   Map<texto, índice>    (shared strings forward map)
 */
function replaceFormulasWithValues(sheetXml, evalMap, ssMap) {
  // Regex que NO hace match de celdas self-closing: el > de apertura NO debe estar precedido de /
  // Usa negative lookbehind (?<!\/) para excluir tags que terminan en />
  // Esto también cubre celdas sin atributos extra: <c r="AF3"> (sin s="N")
  return sheetXml.replace(/<c r="([A-Z]{1,3}\d+)"([^>]*)(?<!\/)>([\s\S]*?)<\/c>/g, (match, ref, attrs, inner) => {
    if (!inner.includes('<f')) return match  // sin fórmula → sin cambios
    let val = evalMap.get(ref)
    // HyperFormula devuelve DetailedCellError para errores (#LIC!, #REF!, etc.) → tratar como vacío
    if (val !== null && val !== undefined && typeof val === 'object' && 'value' in val) val = null
    const baseAttrs = attrs.replace(/\s+t="[^"]*"/, '') // quitar tipo anterior
    if (val === null || val === undefined) {
      return `<c r="${ref}"${baseAttrs}/>`
    }
    if (typeof val === 'number') {
      if (!isFinite(val)) return `<c r="${ref}"${baseAttrs}/>`
      const rounded = Math.round(val * 1e10) / 1e10
      return `<c r="${ref}"${baseAttrs}><v>${rounded}</v></c>`
    }
    if (typeof val === 'boolean') {
      return `<c r="${ref}"${baseAttrs} t="b"><v>${val ? 1 : 0}</v></c>`
    }
    if (typeof val === 'string') {
      if (!val.trim()) return `<c r="${ref}"${baseAttrs}/>`
      const idx = ssMap.get(val)
      if (idx !== undefined) {
        return `<c r="${ref}"${baseAttrs} t="s"><v>${idx}</v></c>`
      }
      // Valor de string no existe en shared strings → inlineStr
      return `<c r="${ref}"${baseAttrs} t="inlineStr"><is><t>${escapeXml(val)}</t></is></c>`
    }
    return `<c r="${ref}"${baseAttrs}/>`
  })
}

/**
 * Agrega a styles.xml: fuente Calibri 9 + fills para cada color único + xf entries.
 * Devuelve { updatedStylesXml, styleIdxMap: Map<argbString, xfIndex> }
 * La clave '' en styleIdxMap corresponde a xf con Calibri 9 sin relleno.
 */
function addStylesForBvac(stylesXml, fillArgbs) {
  let xml = stylesXml

  // 1. Añadir fuente Calibri 9
  const fontCountMatch = xml.match(/<fonts count="(\d+)"/)
  const fontCount = fontCountMatch ? parseInt(fontCountMatch[1]) : 0
  const calibri9Idx = fontCount
  const calibri9Xml = '<font><sz val="9"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>'
  xml = xml.replace(/<\/fonts>/, calibri9Xml + '</fonts>')
  xml = xml.replace(/<fonts count="\d+"/, `<fonts count="${fontCount + 1}"`)

  // 2. Añadir numFmt personalizado DD/MM/YYYY
  const usedNumFmtIds = [...xml.matchAll(/numFmtId="(\d+)"/g)].map(m => parseInt(m[1]))
  const dateNumFmtId = Math.max(163, ...usedNumFmtIds) + 1
  const numFmtXml = `<numFmt numFmtId="${dateNumFmtId}" formatCode="DD/MM/YYYY"/>`
  if (xml.includes('<numFmts')) {
    xml = xml.replace(/<\/numFmts>/, numFmtXml + '</numFmts>')
    xml = xml.replace(/<numFmts count="(\d+)"/, (m, n) => `<numFmts count="${parseInt(n) + 1}"`)
  } else {
    xml = xml.replace(/(<styleSheet[^>]*>)/, `$1<numFmts count="1">${numFmtXml}</numFmts>`)
  }

  // 3. Añadir fills únicos
  const fillIdxMap = new Map() // argb → fillIdx en la tabla
  const fillCountMatch = xml.match(/<fills count="(\d+)"/)
  let fillCount = fillCountMatch ? parseInt(fillCountMatch[1]) : 0
  let newFillsXml = ''
  for (const argb of fillArgbs) {
    if (!argb || fillIdxMap.has(argb)) continue
    fillIdxMap.set(argb, fillCount)
    newFillsXml += `<fill><patternFill patternType="solid"><fgColor rgb="${argb}"/><bgColor indexed="64"/></patternFill></fill>`
    fillCount++
  }
  if (newFillsXml) {
    xml = xml.replace(/<\/fills>/, newFillsXml + '</fills>')
    xml = xml.replace(/<fills count="\d+"/, `<fills count="${fillCount}"`)
  }

  // 4. Añadir xf en cellXfs: regular (numFmtId=0) y fecha por cada color
  const xfCountMatch = xml.match(/<cellXfs count="(\d+)"/)
  let xfCount = xfCountMatch ? parseInt(xfCountMatch[1]) : 0
  const styleIdxMap = new Map()     // argb → xfIndex (texto/número)
  const dateStyleIdxMap = new Map() // argb → xfIndex (fecha DD/MM/YYYY)
  let newXfsXml = ''
  const allArgbs = ['', ...fillArgbs.filter(a => a && !styleIdxMap.has(a))]
  for (const argb of allArgbs) {
    if (styleIdxMap.has(argb)) continue
    const fillId = argb ? fillIdxMap.get(argb) : 0
    const fillAttr = argb ? ` fillId="${fillId}" applyFill="1"` : ' fillId="0"'
    // Estilo regular
    newXfsXml += `<xf numFmtId="0" fontId="${calibri9Idx}" borderId="0" xfId="0"${fillAttr} applyFont="1"/>`
    styleIdxMap.set(argb, xfCount++)
    // Estilo fecha
    newXfsXml += `<xf numFmtId="${dateNumFmtId}" fontId="${calibri9Idx}" borderId="0" xfId="0"${fillAttr} applyFont="1" applyNumberFormat="1"/>`
    dateStyleIdxMap.set(argb, xfCount++)
  }
  xml = xml.replace(/<\/cellXfs>/, newXfsXml + '</cellXfs>')
  xml = xml.replace(/<cellXfs count="\d+"/, `<cellXfs count="${xfCount}"`)

  return { updatedStylesXml: xml, styleIdxMap, dateStyleIdxMap }
}

/**
 * Extrae cualquier hoja como XLSX standalone con fórmulas evaluadas por HyperFormula.
 * Carga todas las hojas dependientes y produce un ZIP mínimo con solo la hoja objetivo.
 */
async function extractSheetStandalone(files, workbookXml, relsXml, sheetName) {
  // ── 1. Shared strings ──────────────────────────────────────────────────────
  const sharedStrXml = strFromU8(files['xl/sharedStrings.xml'])
  const ssArr = buildSsReverseArray(sharedStrXml)
  const { map: ssForwardMap } = parseSharedStrings(sharedStrXml)

  // ── 2. Cargar HyperFormula ─────────────────────────────────────────────────
  const { HyperFormula } = await import('hyperformula')

  // ── 3. Hojas dependientes como valores estáticos (cacheados) ──────────────
  const sheetsData = {}
  const staticSheets = [
    'Parametros', 'Maestro Personal', 'Devengos',
    'rev incapacidades', 'Pensionado', 'Validacion novedades',
    'Planilla Mes anterior', 'Conceptos', 'rete adic', 'revision vac ',
  ]
  for (const name of staticSheets) {
    if (name === sheetName) continue  // la hoja objetivo se carga con fórmulas abajo
    try {
      const p = resolveSheetZipPath(workbookXml, relsXml, name)
      if (p && files[p]) sheetsData[name] = parseSheetXmlToValues(strFromU8(files[p]), ssArr)
    } catch (_) {}
  }

  // Detallado Mes con fórmulas (columnas J-P usan VLOOKUP contra Conceptos)
  if (sheetName !== 'Detallado Mes') {
    try {
      const p = resolveSheetZipPath(workbookXml, relsXml, 'Detallado Mes')
      if (p && files[p]) sheetsData['Detallado Mes'] = parseSheetXmlToFormulas(strFromU8(files[p]), ssArr)
    } catch (_) {}
  }

  // ── 4. Hoja objetivo con fórmulas ──────────────────────────────────────────
  const targetPath = resolveSheetZipPath(workbookXml, relsXml, sheetName)
  const targetXml = strFromU8(files[targetPath])
  sheetsData[sheetName] = parseSheetXmlToFormulas(targetXml, ssArr)

  // ── 5. Instanciar HyperFormula y evaluar ───────────────────────────────────
  const hf = HyperFormula.buildFromSheets(sheetsData, { licenseKey: 'gpl-v3' })
  const targetSheetId = hf.getSheetId(sheetName)
  const targetGrid = sheetsData[sheetName]

  const evalMap = new Map()
  for (let r = 0; r < targetGrid.length; r++) {
    for (let c = 0; c < (targetGrid[r] || []).length; c++) {
      try {
        const val = hf.getCellValue({ sheet: targetSheetId, row: r, col: c })
        if (val !== null && val !== undefined) {
          evalMap.set(`${colToLetter(c + 1)}${r + 1}`, val)
        }
      } catch (_) {}
    }
  }

  // ── 6. Reconstruir XML con valores en lugar de fórmulas ───────────────────
  // Quitar referencia a printerSettings (no existe en el XLSX standalone)
  const cleanedXml = targetXml.replace(/\s+r:id="rId\d+"/g, (m, offset, str) => {
    const before = str.lastIndexOf('<', offset)
    return str.slice(before, offset).includes('pageSetup') ? '' : m
  })
  const evaluatedXml = replaceFormulasWithValues(cleanedXml, evalMap, ssForwardMap)

  // ── 7. Construir ZIP mínimo con solo la hoja objetivo ─────────────────────
  const safeName = sheetName.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')

  const CT_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/></Types>`
  const RELS_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`
  const WB_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="${safeName}" sheetId="1" r:id="rId1"/></sheets></workbook>`
  const WB_RELS_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/></Relationships>`

  const newFiles = {
    '[Content_Types].xml': strToU8(CT_XML),
    '_rels/.rels': strToU8(RELS_XML),
    'xl/workbook.xml': strToU8(WB_XML),
    'xl/_rels/workbook.xml.rels': strToU8(WB_RELS_XML),
    'xl/worksheets/sheet1.xml': strToU8(evaluatedXml),
    'xl/sharedStrings.xml': files['xl/sharedStrings.xml'],
    'xl/styles.xml': files['xl/styles.xml'],
  }
  if (files['xl/theme/theme1.xml']) newFiles['xl/theme/theme1.xml'] = files['xl/theme/theme1.xml']

  return zipSync(newFiles)
}

/** Wrapper: extrae "Seguridad Social" como XLSX standalone */
async function extractSegSocial(files, workbookXml, relsXml) {
  return extractSheetStandalone(files, workbookXml, relsXml, 'Seguridad Social')
}

// ─── Valores por defecto de Parámetros (fallback si la plantilla no carga) ────
const PARAMS_DEFAULTS = {
  uvt: 52374,
  ret: [
    { desde: 95,   tarifa: 0.19, descuento: 0   },
    { desde: 150,  tarifa: 0.28, descuento: 10  },
    { desde: 360,  tarifa: 0.33, descuento: 69  },
    { desde: 640,  tarifa: 0.35, descuento: 162 },
    { desde: 945,  tarifa: 0.37, descuento: 268 },
    { desde: 2300, tarifa: 0.39, descuento: 770 },
  ],
  sol: [
    { rango: 4,  tarifa: 0.01   },
    { rango: 16, tarifa: 0.012  },
    { rango: 17, tarifa: 0.014  },
    { rango: 18, tarifa: 0.016  },
    { rango: 19, tarifa: 0.018  },
    { rango: 20, tarifa: 0.02   },
  ],
  reteAdic: [],
}

/** Lee los valores de parámetros desde la hoja Parametros (sheet10.xml) */
function extractParamsFromSheet(sheetXml) {
  const cellMap = new Map()
  const cellRe = /<c r="([A-Z]{1,3}\d+)"([^>]*)(?<!\/)>([\s\S]*?)<\/c>/g
  let cm
  while ((cm = cellRe.exec(sheetXml)) !== null) {
    const ref = cm[1]
    const inner = cm[3]
    const vM = inner.match(/<v>([^<]*)<\/v>/)
    if (vM) cellMap.set(ref, parseFloat(vM[1]))
  }
  const uvt = cellMap.get('D13') ?? PARAMS_DEFAULTS.uvt
  const ret = []
  const sol = []
  for (let r = 23; r <= 28; r++) {
    ret.push({
      desde:     cellMap.get(`B${r}`) ?? PARAMS_DEFAULTS.ret[r - 23].desde,
      tarifa:    cellMap.get(`C${r}`) ?? PARAMS_DEFAULTS.ret[r - 23].tarifa,
      descuento: cellMap.get(`D${r}`) ?? PARAMS_DEFAULTS.ret[r - 23].descuento,
    })
    sol.push({
      rango:  cellMap.get(`F${r}`) ?? PARAMS_DEFAULTS.sol[r - 23].rango,
      tarifa: cellMap.get(`H${r}`) ?? PARAMS_DEFAULTS.sol[r - 23].tarifa,
    })
  }
  return { uvt, ret, sol, reteAdic: [] }  // reteAdic se carga por separado desde su hoja
}

function downloadBlob(bytes, filename) {
  const blob = new Blob([bytes], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  })
  const url = URL.createObjectURL(blob)
  const a = document.createElement('a')
  a.href = url
  a.download = filename
  document.body.appendChild(a)
  a.click()
  document.body.removeChild(a)
  URL.revokeObjectURL(url)
}

function App() {
  const [file, setFile] = useState(null)
  const [isDragging, setIsDragging] = useState(false)
  const [fileNomina, setFileNomina] = useState(null)
  const [isDraggingNomina, setIsDraggingNomina] = useState(false)
  const [fileMaestro, setFileMaestro] = useState(null)
  const [isDraggingMaestro, setIsDraggingMaestro] = useState(false)
  const [filePlantilla, setFilePlantilla] = useState(null)
  const [isDraggingPlantilla, setIsDraggingPlantilla] = useState(false)
  const [fileNovedades, setFileNovedades] = useState(null)
  const [isDraggingNovedades, setIsDraggingNovedades] = useState(false)
  const [fileNominaAnterior, setFileNominaAnterior] = useState(null)
  const [isDraggingNominaAnterior, setIsDraggingNominaAnterior] = useState(false)
  const [isHelpExpanded, setIsHelpExpanded] = useState(false)
  const [isProcessing, setIsProcessing] = useState(false)
  const [modal, setModal] = useState(null)
  const [params, setParams] = useState(null)
  const [activeTab, setActiveTab] = useState('validacion')
  const fileInputRef = useRef(null)
  const fileNominaRef = useRef(null)
  const fileMaestroRef = useRef(null)
  const filePlantillaRef = useRef(null)
  const fileNovedadesRef = useRef(null)
  const fileNominaAnteriorRef = useRef(null)

  // Función helper: extrae params desde un ArrayBuffer de plantilla
  const loadParamsFromBuffer = (buf) => {
    try {
      const tplFiles = unzipSync(new Uint8Array(buf))
      const wbXml = strFromU8(tplFiles['xl/workbook.xml'])
      const relsXml = strFromU8(tplFiles['xl/_rels/workbook.xml.rels'])
      const ssXml = strFromU8(tplFiles['xl/sharedStrings.xml'])
      const ssArr = buildSsReverseArray(ssXml)
      let loadedParams = { ...PARAMS_DEFAULTS }
      const paramPath = resolveSheetZipPath(wbXml, relsXml, 'Parametros')
      if (paramPath && tplFiles[paramPath]) {
        loadedParams = extractParamsFromSheet(strFromU8(tplFiles[paramPath]))
      }
      const reteAdicPath = resolveSheetZipPath(wbXml, relsXml, 'rete adic')
      if (reteAdicPath && tplFiles[reteAdicPath]) {
        loadedParams.reteAdic = parseReteAdicSheet(strFromU8(tplFiles[reteAdicPath]), ssArr)
      }
      return loadedParams
    } catch {
      return { ...PARAMS_DEFAULTS }
    }
  }

  // Al montar: carga parámetros desde la plantilla por defecto (fallback)
  useEffect(() => {
    fetch('/Ejemplo Validacion de Nomina.xlsx')
      .then(r => r.ok ? r.arrayBuffer() : Promise.reject())
      .then(buf => setParams(loadParamsFromBuffer(buf)))
      .catch(() => setParams({ ...PARAMS_DEFAULTS }))
  }, [])

  const validateExcel = (f) => {
    if (!f) return false
    if (!f.name.match(/\.xlsx?$/i)) {
      setModal({ type: 'error', message: 'Por favor sube un archivo Excel (.xlsx o .xls)' })
      return false
    }
    return true
  }

  const handleFile = (f) => { if (validateExcel(f)) setFile(f) }
  const handleFileNomina = (f) => { if (validateExcel(f)) setFileNomina(f) }
  const handleFileMaestro = (f) => { if (validateExcel(f)) setFileMaestro(f) }
  const handleFilePlantilla = (f) => {
    if (!validateExcel(f)) return
    setFilePlantilla(f)
    // Recargar parámetros desde la nueva plantilla
    f.arrayBuffer().then(buf => setParams(loadParamsFromBuffer(buf))).catch(() => {})
  }

  const handleDrop = (e) => {
    e.preventDefault()
    setIsDragging(false)
    handleFile(e.dataTransfer.files[0])
  }

  const handleDropNomina = (e) => {
    e.preventDefault()
    setIsDraggingNomina(false)
    handleFileNomina(e.dataTransfer.files[0])
  }

  const handleDropMaestro = (e) => {
    e.preventDefault()
    setIsDraggingMaestro(false)
    handleFileMaestro(e.dataTransfer.files[0])
  }

  const handleDropPlantilla = (e) => {
    e.preventDefault()
    setIsDraggingPlantilla(false)
    handleFilePlantilla(e.dataTransfer.files[0])
  }

  const handleFileNovedades = (f) => { if (validateExcel(f)) setFileNovedades(f) }
  const handleDropNovedades = (e) => {
    e.preventDefault()
    setIsDraggingNovedades(false)
    handleFileNovedades(e.dataTransfer.files[0])
  }

  const handleFileNominaAnterior = (f) => { if (validateExcel(f)) setFileNominaAnterior(f) }
  const handleDropNominaAnterior = (e) => {
    e.preventDefault()
    setIsDraggingNominaAnterior(false)
    handleFileNominaAnterior(e.dataTransfer.files[0])
  }

  const removeFile = () => {
    setFile(null)
    if (fileInputRef.current) fileInputRef.current.value = ''
  }

  const removeFileNomina = () => {
    setFileNomina(null)
    if (fileNominaRef.current) fileNominaRef.current.value = ''
  }

  const removeFileMaestro = () => {
    setFileMaestro(null)
    if (fileMaestroRef.current) fileMaestroRef.current.value = ''
  }

  const removeFilePlantilla = () => {
    setFilePlantilla(null)
    if (filePlantillaRef.current) filePlantillaRef.current.value = ''
    // Volver a cargar params desde plantilla por defecto
    fetch('/Ejemplo Validacion de Nomina.xlsx')
      .then(r => r.ok ? r.arrayBuffer() : Promise.reject())
      .then(buf => setParams(loadParamsFromBuffer(buf)))
      .catch(() => setParams({ ...PARAMS_DEFAULTS }))
  }

  const removeFileNovedades = () => {
    setFileNovedades(null)
    if (fileNovedadesRef.current) fileNovedadesRef.current.value = ''
  }

  const removeFileNominaAnterior = () => {
    setFileNominaAnterior(null)
    if (fileNominaAnteriorRef.current) fileNominaAnteriorRef.current.value = ''
  }

  const formatBytes = (bytes) => {
    if (bytes < 1024) return bytes + ' B'
    if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB'
    return (bytes / (1024 * 1024)).toFixed(1) + ' MB'
  }

  const processFile = async () => {
    if (!file && !fileNomina && !fileMaestro) return
    setIsProcessing(true)
    try {
      let dataRows = []
      let nominaRows = []
      let maestroRows = []  // filas filtradas del Maestro Personal

      // ── Leer IBC si fue cargado ──────────────────────────────────────────────
      if (file) {
        const ibcBuffer = await file.arrayBuffer()
        const ibcWb = new ExcelJS.Workbook()
        await ibcWb.xlsx.load(ibcBuffer)
        const ibcSheet = ibcWb.getWorksheet('Sheet1')
        if (!ibcSheet) throw new Error('No se encontró la hoja "Sheet1" en el archivo IBC.')

        let currentCity = ''
        let currentDepto = ''
        ibcSheet.eachRow({ includeEmpty: false }, (row, rowNum) => {
          if (rowNum < 20) return
          const colBVal = getCellValue(row.getCell(2))
          const isNumeric = colBVal !== null && colBVal !== '' && !isNaN(Number(colBVal))
          if (!isNumeric) {
            let rowStr = ''
            row.eachCell({ includeEmpty: false }, (cell) => {
              const v = getCellValue(cell)
              if (v !== null && v !== undefined) rowStr += ' ' + String(v)
            })
            const dm = rowStr.match(/Ciudad:\s*(.+?)\s+Depto:\s*([^(,\n]+)/i)
            if (dm) { currentCity = dm[1].trim(); currentDepto = dm[2].trim() }
          } else {
            dataRows.push({ row, city: currentCity, depto: currentDepto })
          }
        })
        if (dataRows.length === 0) throw new Error('No se encontraron registros válidos en el archivo IBC.')
      }

      // ── Leer Nómina si fue cargada ───────────────────────────────────────────
      if (fileNomina) {
        const nomBuffer = await fileNomina.arrayBuffer()
        const nomWb = new ExcelJS.Workbook()
        await nomWb.xlsx.load(nomBuffer)
        const nomSheet = nomWb.worksheets[0]
        if (!nomSheet) throw new Error('No se encontró ninguna hoja en el archivo de nómina.')
        nomSheet.eachRow({ includeEmpty: false }, (row, rowNum) => {
          if (rowNum <= 4) return // encabezados en fila 4
          nominaRows.push(row)
        })
        if (nominaRows.length === 0) throw new Error('No se encontraron registros en el archivo de nómina.')
      }

      // ── Leer Maestro Personal si fue cargado ─────────────────────────────────
      if (fileMaestro) {
        const mpBuffer = await fileMaestro.arrayBuffer()
        const mpWb = new ExcelJS.Workbook()
        await mpWb.xlsx.load(mpBuffer)
        const mpSheet = mpWb.worksheets[0]
        if (!mpSheet) throw new Error('No se encontró ninguna hoja en el archivo Maestro Personal.')
        const ESTADOS_VALIDOS = new Set(['Activos', 'Ausencia', 'Vacaciones'])
        mpSheet.eachRow({ includeEmpty: false }, (row, rowNum) => {
          if (rowNum <= 5) return // headers en fila 3, ignorar filas 1-5
          const estado = String(getCellValue(row.getCell(61)) ?? '').trim()
          if (ESTADOS_VALIDOS.has(estado)) maestroRows.push(row)
        })
        if (maestroRows.length === 0) throw new Error('No se encontraron registros Activos/Ausencia/Vacaciones en el Maestro Personal.')
      }

      // ── Cargar plantilla ─────────────────────────────────────────────────────
      let templateBytes
      if (filePlantilla) {
        templateBytes = new Uint8Array(await filePlantilla.arrayBuffer())
      } else {
        const res = await fetch('/Ejemplo Validacion de Nomina.xlsx')
        if (!res.ok) throw new Error('No se pudo cargar la plantilla de validación.')
        templateBytes = new Uint8Array(await res.arrayBuffer())
      }
      const files = unzipSync(templateBytes)
      const workbookXml = strFromU8(files['xl/workbook.xml'])
      const relsXml = strFromU8(files['xl/_rels/workbook.xml.rels'])

      // ── Recopilar strings necesarios (ambas hojas) ───────────────────────────
      const NO_DEFAULT_COLS = new Set([13, 14, 16, 18, 20, 21, 22, 23, 24, 26, 27, 30, 33, 36, 39, 40])
      const neededStrings = new Set(['SI', 'NO',
        'Todos los sistemas (ARL, AFP, CCF, EPS)', 'INCAPACIDAD GENERAL',
        'VACACIONES', 'LICENCIA REMUNERADA'])

      if (file) {
        for (const { row: srcRow, city, depto } of dataRows) {
          for (const { src, dst } of FIELD_MAP) {
            const v = transformValue(getCellValue(srcRow.getCell(src)), dst)
            if (v !== null && v !== undefined && typeof v !== 'number' && typeof v !== 'boolean') {
              const str = v instanceof Date ? formatDate(v) : String(v).trim()
              if (str) neededStrings.add(str)
            }
          }
          const nombre = getCellValue(srcRow.getCell(10))
          if (nombre) String(nombre).trim().split(/\s+/).filter(Boolean).forEach(p => neededStrings.add(p))
          if (city) neededStrings.add(city)
          if (depto) neededStrings.add(depto)
        }
      }

      if (fileNomina) {
        for (const srcRow of nominaRows) {
          // Cols de texto para Detallado Mes (1-4) + cols adicionales para BVAC (10, 13-16, 18)
          for (const col of [1, 2, 3, 4, 10, 13, 14, 15, 16, 18]) {
            const v = getCellValue(srcRow.getCell(col))
            if (v != null && !(v instanceof Date)) { const s = String(v).trim(); if (s) neededStrings.add(s) }
          }
        }
      }

      if (fileMaestro) {
        for (const srcRow of maestroRows) {
          for (let col = 2; col <= 61; col++) {
            const spec = MAESTRO_COL_SPEC[col] || { type: 'str', style: 83 }
            if (spec.type === 'str' || spec.type === 'num') {
              const v = getCellValue(srcRow.getCell(col))
              if (v != null && !(v instanceof Date)) {
                const n = typeof v === 'number' ? v : Number(v)
                if (spec.type === 'str' || isNaN(n)) {
                  const s = String(v).trim()
                  if (s) neededStrings.add(s)
                }
              }
            }
          }
        }
      }

      // ── Actualizar sharedStrings.xml ─────────────────────────────────────────
      const ssXml = strFromU8(files['xl/sharedStrings.xml'])
      const { strIdxMap, newSsXml } = buildSharedStringIndex(ssXml, neededStrings)
      files['xl/sharedStrings.xml'] = strToU8(newSsXml)

      // ── Procesar Planilla Mes anterior ───────────────────────────────────────
      if (file) {
        const sheetPath = resolveSheetZipPath(workbookXml, relsXml, 'Planilla Mes anterior')
        if (!files[sheetPath]) throw new Error(`No se encontró el archivo de hoja: ${sheetPath}`)

        // Calcular resultados de la tabla dinámica TablaDinámica1 (CX:CY):
        // Agrupa por "No ID" (src=8) y suma "IBC2" (src=86, col BM)
        const ibcPivotMap = new Map()  // key=string → { id: valor original, sum: número }
        for (const { row: srcRow } of dataRows) {
          const noId = getCellValue(srcRow.getCell(8))
          const ibc2Raw = getCellValue(srcRow.getCell(86))
          if (noId != null && String(noId).trim()) {
            const key = String(noId).trim()
            const val = typeof ibc2Raw === 'number' ? ibc2Raw : (parseFloat(ibc2Raw) || 0)
            if (!ibcPivotMap.has(key)) ibcPivotMap.set(key, { id: noId, sum: 0 })
            ibcPivotMap.get(key).sum += val
          }
        }
        // Ordenar igual que la tabla dinámica (por No ID ascendente como string)
        const pivotRows = [...ibcPivotMap.entries()]
          .sort((a, b) => String(a[0]).localeCompare(String(b[0])))
          .map(([, { id, sum }]) => [id, sum])  // preserva tipo original de id (número)

        const newRowsXml = dataRows.map(({ row: srcRow, city, depto }, idx) => {
          const rowNum = idx + 2
          const valMap = new Map()
          for (const { src, dst } of FIELD_MAP) {
            valMap.set(dst, transformValue(getCellValue(srcRow.getCell(src)), dst))
          }
          const nameParts = splitNombre(getCellValue(srcRow.getCell(10)))
          for (const [dstCol, part] of Object.entries(nameParts)) valMap.set(Number(dstCol), part)
          if (depto) valMap.set(8, depto)
          if (city) valMap.set(9, city)
          for (const dst of NO_DEFAULT_COLS) {
            const v = valMap.get(dst)
            const isEmpty = v === null || v === undefined || v === 0 || v === false ||
              (typeof v === 'string' && !v.trim())
            if (isEmpty) valMap.set(dst, 'NO')
          }
          // Inyectar resultado de tabla dinámica en CX(102) y CY(103)
          if (idx < pivotRows.length) {
            const [pNoId, pIbc2] = pivotRows[idx]
            valMap.set(102, pNoId)   // CX = "No ID" (etiqueta de fila)
            valMap.set(103, pIbc2)  // CY = "Suma de IBC2"
          }
          const cells = [...valMap.entries()]
            .sort((a, b) => a[0] - b[0])
            .map(([dst, v]) => buildCell(dst, rowNum, v, strIdxMap))
            .filter(Boolean).join('')
          return cells
            ? `<row r="${rowNum}" spans="1:104" s="55" customFormat="1" x14ac:dyDescent="0.3">${cells}</row>`
            : ''
        }).filter(Boolean).join('')

        files[sheetPath] = strToU8(rebuildSheetData(strFromU8(files[sheetPath]), newRowsXml, 1))
      }

      // ── Procesar Detallado Mes ───────────────────────────────────────────────
      if (fileNomina) {
        const detPath = resolveSheetZipPath(workbookXml, relsXml, 'Detallado Mes')
        if (!files[detPath]) throw new Error('No se encontró la hoja "Detallado Mes" en la plantilla.')

        const detRowsXml = nominaRows.map((srcRow, idx) => {
          const r = idx + 4 // datos desde fila 4 (encabezados en fila 3)
          const descConceptoVal = getCellValue(srcRow.getCell(4)) // NOMBRE CONCEPTO (col E en Detallado Mes)
          const cantidadVal = String(descConceptoVal ?? '').trim() === 'SIP Payment'
            ? 0
            : getCellValue(srcRow.getCell(5)) // CANTIDAD → HORAS
          return buildDetalladoRow(
            r,
            getCellValue(srcRow.getCell(1)), // CODIGO EMPLEADO
            getCellValue(srcRow.getCell(2)), // NOMBRES EMPLEADO
            getCellValue(srcRow.getCell(3)), // CONCEPTO
            descConceptoVal,                 // NOMBRE CONCEPTO
            cantidadVal,                     // CANTIDAD (0 si SIP Payment)
            getCellValue(srcRow.getCell(6)), // DEVENGOS
            getCellValue(srcRow.getCell(7)), // DEDUCCIONES
            strIdxMap
          )
        }).join('')

        files[detPath] = strToU8(rebuildSheetData(strFromU8(files[detPath]), detRowsXml, 3))

        // ── Anexar nómina a BVAC (debajo del último registro existente) ──────
        try {
          const bvacPath = resolveSheetZipPath(workbookXml, relsXml, 'BVAC ')
          if (bvacPath && files[bvacPath]) {
            // Recoger colores de relleno únicos desde las filas de nómina
            const uniqueFillArgbs = []
            const seenArgbs = new Set()
            for (const srcRow of nominaRows) {
              const fill = srcRow.fill ?? srcRow.getCell?.(1)?.fill
              const argb = fill?.fgColor?.argb ?? ''
              if (!seenArgbs.has(argb)) { seenArgbs.add(argb); uniqueFillArgbs.push(argb) }
            }
            // Añadir Calibri 9 + fills + estilos fecha a styles.xml y obtener mapas de estilos
            const { updatedStylesXml: bvacStylesXml, styleIdxMap: bvacStyleMap, dateStyleIdxMap: bvacDateStyleMap } =
              addStylesForBvac(strFromU8(files['xl/styles.xml']), uniqueFillArgbs)
            files['xl/styles.xml'] = strToU8(bvacStylesXml)

            const bvacXml = strFromU8(files[bvacPath])
            // Encontrar la última fila con datos
            const rowNums = [...bvacXml.matchAll(/<row r="(\d+)"/g)].map(m => parseInt(m[1]))
            const lastRow = rowNums.length > 0 ? Math.max(...rowNums) : 1

            const bvacNewRows = nominaRows.map((srcRow, idx) => {
              const r = lastRow + 1 + idx
              // Estilo: Calibri 9 + color de fila
              const fill = srcRow.fill ?? srcRow.getCell?.(1)?.fill
              const fillArgb = fill?.fgColor?.argb ?? ''
              const sIdx = bvacStyleMap.get(fillArgb) ?? bvacStyleMap.get('')
              const s = sIdx !== undefined ? ` s="${sIdx}"` : ''
              const dsIdx = bvacDateStyleMap.get(fillArgb) ?? bvacDateStyleMap.get('')
              const sDate = dsIdx !== undefined ? ` s="${dsIdx}"` : s

              let c = ''
              // A: CODIGO EMPLEADO (numérico)
              const emp = getCellValue(srcRow.getCell(1))
              const empStr = (emp == null ? '' : String(emp)).trim()
              if (empStr && /^\d+$/.test(empStr)) c += `<c r="A${r}"${s}><v>${empStr}</v></c>`
              // B: NOMBRES EMPLEADO (texto)
              const bIdx = strIdxMap.get(String(getCellValue(srcRow.getCell(2)) ?? '').trim())
              if (bIdx !== undefined) c += `<c r="B${r}"${s} t="s"><v>${bIdx}</v></c>`
              // C: CONCEPTO (texto — preservar ceros iniciales como "001050")
              const cIdx = strIdxMap.get(String(getCellValue(srcRow.getCell(3)) ?? '').trim())
              if (cIdx !== undefined) c += `<c r="C${r}"${s} t="s"><v>${cIdx}</v></c>`
              // D: NOMBRE CONCEPTO (texto)
              const dIdx = strIdxMap.get(String(getCellValue(srcRow.getCell(4)) ?? '').trim())
              if (dIdx !== undefined) c += `<c r="D${r}"${s} t="s"><v>${dIdx}</v></c>`
              // E: CANTIDAD (número)
              const cant = Number(getCellValue(srcRow.getCell(5)) ?? 0)
              c += `<c r="E${r}"${s}><v>${isNaN(cant) ? 0 : cant}</v></c>`
              // F: DEVENGOS, G: DEDUCCIONES, H: PAGOS INDIRECTOS, I: NETO PAGADO
              for (const [colNum, letter] of [[6,'F'],[7,'G'],[8,'H'],[9,'I']]) {
                const n = Number(getCellValue(srcRow.getCell(colNum)) ?? 0)
                c += `<c r="${letter}${r}"${s}><v>${isNaN(n) ? 0 : n}</v></c>`
              }
              // J: TIPO DE LIQUIDACION (texto)
              const jIdx = strIdxMap.get(String(getCellValue(srcRow.getCell(10)) ?? '').trim())
              if (jIdx !== undefined) c += `<c r="J${r}"${s} t="s"><v>${jIdx}</v></c>`
              // K: FECHA DE CORTE, L: FECHA DE LIQUIDACION (fecha → serial Excel con formato DD/MM/YYYY)
              for (const [colNum, letter] of [[11,'K'],[12,'L']]) {
                const raw = getCellValue(srcRow.getCell(colNum))
                if (raw instanceof Date) {
                  c += `<c r="${letter}${r}"${sDate}><v>${jsDateToExcelSerial(raw)}</v></c>`
                } else if (raw != null && !isNaN(Number(raw))) {
                  c += `<c r="${letter}${r}"${sDate}><v>${Number(raw)}</v></c>`
                }
              }
              // M: ORIGEN, N: COMPAÑIA, O: SUCURSAL (texto)
              for (const [colNum, letter] of [[13,'M'],[14,'N'],[15,'O']]) {
                const sidx = strIdxMap.get(String(getCellValue(srcRow.getCell(colNum)) ?? '').trim())
                if (sidx !== undefined) c += `<c r="${letter}${r}"${s} t="s"><v>${sidx}</v></c>`
              }
              // P: CENTRO DE COSTOS (texto — ej. "5180")
              const pIdx = strIdxMap.get(String(getCellValue(srcRow.getCell(16)) ?? '').trim())
              if (pIdx !== undefined) c += `<c r="P${r}"${s} t="s"><v>${pIdx}</v></c>`
              // Q: SUELDO BASICO (número)
              const q = Number(getCellValue(srcRow.getCell(17)) ?? 0)
              if (!isNaN(q)) c += `<c r="Q${r}"${s}><v>${q}</v></c>`
              // R: CARGO (texto)
              const rIdx = strIdxMap.get(String(getCellValue(srcRow.getCell(18)) ?? '').trim())
              if (rIdx !== undefined) c += `<c r="R${r}"${s} t="s"><v>${rIdx}</v></c>`

              return c ? `<row r="${r}" spans="1:18" x14ac:dyDescent="0.3">${c}</row>` : ''
            }).filter(Boolean).join('')

            const newLastRow = lastRow + nominaRows.length
            const bvacUpdated = bvacXml
              .replace(/<dimension ref="[^"]*"\/>/,  `<dimension ref="A1:R${newLastRow}"/>`)
              .replace(/<\/sheetData>/, bvacNewRows + '</sheetData>')
            files[bvacPath] = strToU8(bvacUpdated)
          }
        } catch (_) { /* BVAC no encontrado, continuar */ }

        // ── Procesar rev incapacidades ───────────────────────────────────────
        const INCAP_CONCEPTOS = new Set(['001150', '001151', '001177', '001178'])
        const incapRows = nominaRows.filter(srcRow => {
          const concepto = String(getCellValue(srcRow.getCell(3)) ?? '').trim()
          return INCAP_CONCEPTOS.has(concepto)
        })
        const revPath = resolveSheetZipPath(workbookXml, relsXml, 'rev incapacidades')
        if (files[revPath] && incapRows.length > 0) {
          // Pre-construir mapa empCode → cantidad del concepto 001150 (para usarla en 001177)
          const emp001150CantMap = new Map()
          for (const srcRow of incapRows) {
            const conc = String(getCellValue(srcRow.getCell(3)) ?? '').trim()
            if (conc === '001150') {
              const emp = String(getCellValue(srcRow.getCell(1)) ?? '').trim()
              const cant = Number(getCellValue(srcRow.getCell(5)) ?? 0)
              if (emp) emp001150CantMap.set(emp, isNaN(cant) ? 0 : cant)
            }
          }
          const revRowsXml = incapRows.map((srcRow, idx) => {
            const r = idx + 13 // datos desde fila 13 (encabezados en fila 12)
            const conc = String(getCellValue(srcRow.getCell(3)) ?? '').trim()
            let kRef = '$C$4'
            let cantidadOverride = null
            let pct = '70%'
            if (conc === '001150') {
              kRef = '$C$3'
            } else if (conc === '001177') {
              kRef = '$C$5'
              const emp = String(getCellValue(srcRow.getCell(1)) ?? '').trim()
              if (emp001150CantMap.has(emp)) cantidadOverride = emp001150CantMap.get(emp)
            } else if (conc === '001178') {
              kRef = '$C$6'
              pct = '100%'
            }
            return buildRevIncapRow(
              r,
              getCellValue(srcRow.getCell(1)), // CODIGO EMPLEADO
              getCellValue(srcRow.getCell(2)), // NOMBRES EMPLEADO
              getCellValue(srcRow.getCell(3)), // CONCEPTO
              getCellValue(srcRow.getCell(4)), // NOMBRE CONCEPTO
              getCellValue(srcRow.getCell(5)), // CANTIDAD
              getCellValue(srcRow.getCell(6)), // DEVENGOS
              getCellValue(srcRow.getCell(7)), // DEDUCCIONES
              strIdxMap,
              kRef,
              cantidadOverride,
              pct
            )
          }).join('')
          files[revPath] = strToU8(rebuildSheetData(strFromU8(files[revPath]), revRowsXml, 12))

          /* ── TablaDinámica4: calcular pivot P:S en "rev incapacidades" ──────
          // COMENTADO TEMPORALMENTE PARA VERIFICACIÓN
          // Fuente: B12:G204 misma hoja. Filas=CODIGO EMPLEADO, Cols=NOMBRE CONCEPTO
          // Col Q = "Inc. Enfermedad Comun Asumida", Col R = "Inc. por Enfermedad Comun"
          // Datos = SUM CANTIDAD (col F de la fuente = col 5 de srcRow nómina)
          const pivot4Map = new Map() // key=empId str → { id: original, q: 0, r: 0 }
          for (const srcRow of incapRows) {
            const emp = getCellValue(srcRow.getCell(1))
            const empKey = String(emp ?? '').trim()
            const nombreConc = String(getCellValue(srcRow.getCell(4)) ?? '').trim().toLowerCase()
            const cantidad = Number(getCellValue(srcRow.getCell(5)) ?? 0) || 0
            if (!empKey) continue
            if (!pivot4Map.has(empKey)) pivot4Map.set(empKey, { id: emp, q: 0, r: 0 })
            const entry = pivot4Map.get(empKey)
            if (nombreConc.includes('inc. enfermedad comun asumida')) entry.q += cantidad
            else if (nombreConc.startsWith('inc. por enfermedad comun')) entry.r += cantidad
          }
          const pivot4Sorted = [...pivot4Map.entries()]
            .sort((a, b) => String(a[0]).localeCompare(String(b[0])))
            .map(([, v]) => v)

          const pivot4CellMap = new Map() // rowNum → cellsXml
          pivot4Sorted.forEach(({ id, q, r: rVal }, idx) => {
            const rowNum = 3 + idx
            const empStr = String(id ?? '').trim()
            const empIsNum = empStr !== '' && /^\d+$/.test(empStr)
            let cells = ''
            if (empIsNum) {
              cells += `<c r="P${rowNum}"><v>${empStr}</v></c>`
            } else {
              const si = strIdxMap.get(empStr)
              if (si !== undefined) cells += `<c r="P${rowNum}" t="s"><v>${si}</v></c>`
            }
            if (q !== 0) cells += `<c r="Q${rowNum}"><v>${q}</v></c>`
            if (rVal !== 0) cells += `<c r="R${rowNum}"><v>${rVal}</v></c>`
            const s = q + rVal
            if (s !== 0) cells += `<c r="S${rowNum}"><v>${s}</v></c>`
            if (cells) pivot4CellMap.set(rowNum, cells)
          })
          // Fila "Total general"
          if (pivot4Sorted.length > 0) {
            const gtRow = 3 + pivot4Sorted.length
            const totalQ = pivot4Sorted.reduce((a, v) => a + v.q, 0)
            const totalR = pivot4Sorted.reduce((a, v) => a + v.r, 0)
            const totalS = totalQ + totalR
            const gtSi = strIdxMap.get('Total general')
            let cells = gtSi !== undefined
              ? `<c r="P${gtRow}" t="s"><v>${gtSi}</v></c>`
              : `<c r="P${gtRow}" t="inlineStr"><is><t>Total general</t></is></c>`
            if (totalQ !== 0) cells += `<c r="Q${gtRow}"><v>${totalQ}</v></c>`
            if (totalR !== 0) cells += `<c r="R${gtRow}"><v>${totalR}</v></c>`
            if (totalS !== 0) cells += `<c r="S${gtRow}"><v>${totalS}</v></c>`
            pivot4CellMap.set(gtRow, cells)
          }
          // Inyectar pivot en sheet (limpia celdas P-S antiguas e inserta las nuevas)
          let revSheetXml = strFromU8(files[revPath])
          revSheetXml = overwritePivotInSheet(revSheetXml, pivot4CellMap)
          files[revPath] = strToU8(revSheetXml)
          ── FIN TablaDinámica4 comentado */
        }

        // ── Procesar revision vac ────────────────────────────────────────────
        const VAC_CONCEPTOS = new Set(['Vacaciones', 'Ajuste Vacaciones', 'Vacaciones en Dinero'])
        const vacRows = nominaRows.filter(srcRow => {
          const desc = String(getCellValue(srcRow.getCell(4)) ?? '').trim()
          return VAC_CONCEPTOS.has(desc)
        })
        const rvPath = resolveSheetZipPath(workbookXml, relsXml, 'revision vac ')
        if (rvPath && files[rvPath] && vacRows.length > 0) {
          const rvRowsXml = vacRows.map((srcRow, idx) => {
            const r = idx + 3 // datos desde fila 3 (encabezados en fila 2)
            const getIdx = (v) => {
              const s = (v == null ? '' : String(v)).trim()
              return s ? strIdxMap.get(s) : undefined
            }
            let c = ''
            // A: código empleado (numérico, s="84")
            const empStr = (getCellValue(srcRow.getCell(1)) == null ? '' : String(getCellValue(srcRow.getCell(1)))).trim()
            if (empStr && /^\d+$/.test(empStr)) {
              c += `<c r="A${r}" s="84"><v>${empStr}</v></c>`
            } else if (empStr) {
              const si = getIdx(empStr)
              if (si !== undefined) c += `<c r="A${r}" s="84" t="s"><v>${si}</v></c>`
            }
            // B: nombre empleado (shared string, s="84")
            const bSi = getIdx(getCellValue(srcRow.getCell(2)))
            if (bSi !== undefined) c += `<c r="B${r}" s="84" t="s"><v>${bSi}</v></c>`
            // C: código concepto (shared string, s="84")
            const cSi = getIdx(getCellValue(srcRow.getCell(3)))
            if (cSi !== undefined) c += `<c r="C${r}" s="84" t="s"><v>${cSi}</v></c>`
            // D: nombre concepto (shared string, s="84")
            const dSi = getIdx(getCellValue(srcRow.getCell(4)))
            if (dSi !== undefined) c += `<c r="D${r}" s="84" t="s"><v>${dSi}</v></c>`
            // E: cantidad (numérico, s="85")
            const cant = Number(getCellValue(srcRow.getCell(5)) ?? 0)
            c += `<c r="E${r}" s="85"><v>${isNaN(cant) ? 0 : cant}</v></c>`
            // F: devengos (numérico, s="85")
            const dev = Number(getCellValue(srcRow.getCell(6)) ?? 0)
            c += `<c r="F${r}" s="85"><v>${isNaN(dev) ? 0 : dev}</v></c>`
            // H-N: fórmulas (ajustadas a la fila r)
            c += `<c r="H${r}"><f>VLOOKUP(A${r},'Maestro Personal'!$B$2:$AH$333,33,0)</f></c>`
            c += `<c r="I${r}" s="22"><f>VLOOKUP(A${r},'Maestro Personal'!$B$2:$AA$333,26,0)</f></c>`
            c += `<c r="J${r}" s="189"><f>+IF(DAYS360(I${r},$J$1)+1&gt;360,360,DAYS360(I${r},$J$1)+3)</f></c>`
            c += `<c r="K${r}" s="60"><f>IFERROR(VLOOKUP(A${r},'BVAC '!$Y:$Z,2,0),0)/J${r}*30</f></c>`
            c += `<c r="L${r}" s="60"><f>+K${r}+H${r}</f></c>`
            c += `<c r="M${r}" s="60"><f>L${r}/30*E${r}</f></c>`
            c += `<c r="N${r}" s="158"><f>M${r}-F${r}</f></c>`
            return `<row r="${r}" spans="1:15" x14ac:dyDescent="0.3">${c}</row>`
          }).join('')
          const rvLastRow = 2 + vacRows.length
          let rvXml = strFromU8(files[rvPath])
          rvXml = rvXml.replace(/<dimension ref="[^"]*"\/>/, `<dimension ref="A1:O${rvLastRow}"/>`)
          rvXml = rebuildSheetData(rvXml, rvRowsXml, 2)
          files[rvPath] = strToU8(rvXml)
        }
      }

      // ── Procesar Maestro Personal ────────────────────────────────────────────
      if (fileMaestro && maestroRows.length > 0) {
        const mpPath = resolveSheetZipPath(workbookXml, relsXml, 'Maestro Personal')
        if (!files[mpPath]) throw new Error('No se encontró la hoja "Maestro Personal" en la plantilla.')
        const mpRowsXml = maestroRows.map((srcRow, idx) => {
          const r = idx + 2  // datos desde fila 2 (encabezados en fila 1)
          return buildMaestroRow(r, srcRow, idx + 1, strIdxMap)
        }).join('')
        files[mpPath] = strToU8(rebuildSheetData(strFromU8(files[mpPath]), mpRowsXml, 1))

        // ── Procesar validacion aux alimentacion (2) ─────────────────────────
        // Construir mapa empCode → concepto (111500 o 111501) desde nómina mes anterior
        const auxAlimPrevMap = new Map()
        if (fileNominaAnterior) {
          const prevBuf = await fileNominaAnterior.arrayBuffer()
          const prevWb = new ExcelJS.Workbook()
          await prevWb.xlsx.load(prevBuf)
          const prevSheet = prevWb.worksheets[0]
          if (prevSheet) {
            prevSheet.eachRow({ includeEmpty: false }, (row, rowNum) => {
              if (rowNum < 2) return
              const emp = String(getCellValue(row.getCell(1)) ?? '').trim()
              const conc = String(getCellValue(row.getCell(3)) ?? '').trim()
              if (emp && (conc === '111501' || conc === '111500')) {
                auxAlimPrevMap.set(emp, conc)
              }
            })
          }
        }
        const auxAlimPath = resolveSheetZipPath(workbookXml, relsXml, 'validacion aux alimentacion (2)')
        if (files[auxAlimPath]) {
          const auxAlimRowsXml = maestroRows.map((srcRow, idx) => {
            const r = idx + 5 // datos desde fila 5
            let c = ''
            // A: CODIGO EMPLEADO (col B del maestro = col 2)
            const empRaw = getCellValue(srcRow.getCell(2))
            const empStr = (empRaw == null ? '' : String(empRaw)).trim()
            const empIsNum = empStr !== '' && /^\d+$/.test(empStr)
            if (empIsNum) {
              c += `<c r="A${r}" s="222"><v>${empStr}</v></c>`
            } else if (empStr) {
              const si = strIdxMap.get(empStr)
              if (si !== undefined) c += `<c r="A${r}" s="222" t="s"><v>${si}</v></c>`
            }
            // B: DOCUMENTO (col E = col 5)
            const docRaw = getCellValue(srcRow.getCell(5))
            const docStr = (docRaw == null ? '' : String(docRaw)).trim()
            if (docStr) {
              const si = strIdxMap.get(docStr)
              if (si !== undefined) c += `<c r="B${r}" s="223" t="s"><v>${si}</v></c>`
              else c += `<c r="B${r}" s="223" t="inlineStr"><is><t>${escapeXml(docStr)}</t></is></c>`
            }
            // C: NOMBRE (col F = col 6)
            const nomRaw = getCellValue(srcRow.getCell(6))
            const nomStr = (nomRaw == null ? '' : String(nomRaw)).trim()
            if (nomStr) {
              const si = strIdxMap.get(nomStr)
              if (si !== undefined) c += `<c r="C${r}" s="223" t="s"><v>${si}</v></c>`
              else c += `<c r="C${r}" s="223" t="inlineStr"><is><t>${escapeXml(nomStr)}</t></is></c>`
            }
            // D: CLASE DE SALARIO (col AG = col 33)
            const claseRaw = getCellValue(srcRow.getCell(33))
            const claseStr = (claseRaw == null ? '' : String(claseRaw)).trim()
            if (claseStr) {
              const si = strIdxMap.get(claseStr)
              if (si !== undefined) c += `<c r="D${r}" s="223" t="s"><v>${si}</v></c>`
              else c += `<c r="D${r}" s="223" t="inlineStr"><is><t>${escapeXml(claseStr)}</t></is></c>`
            }
            // E: SUELDO (col AH = col 34)
            const sueldoRaw = getCellValue(srcRow.getCell(34))
            const sueldoNum = typeof sueldoRaw === 'number' ? sueldoRaw : Number(sueldoRaw)
            if (sueldoRaw != null && String(sueldoRaw).trim() !== '' && !isNaN(sueldoNum)) {
              c += `<c r="E${r}" s="224"><v>${sueldoNum}</v></c>`
            }
            // F: =+E{r}/$F$2
            c += `<c r="F${r}" s="225"><f>+E${r}/$F$2</f></c>`
            // G: =IF(F{r}>=$H$1,"111501","111500")
            c += `<c r="G${r}" s="209" t="str"><f>IF(F${r}&gt;=$H$1,"111501","111500")</f></c>`
            // H: concepto mes anterior si se subió el archivo y se encontró; si no, VLOOKUP
            const prevConc = auxAlimPrevMap.get(empStr)
            if (fileNominaAnterior && prevConc) {
              const si = strIdxMap.get(prevConc)
              if (si !== undefined) c += `<c r="H${r}" s="211" t="s"><v>${si}</v></c>`
              else c += `<c r="H${r}" s="211" t="inlineStr"><is><t>${escapeXml(prevConc)}</t></is></c>`
            } else if (!fileNominaAnterior) {
              c += `<c r="H${r}" s="211" t="str"><f>VLOOKUP(A${r},'[5]a dic va'!$A:$C,3,0)</f></c>`
            }
            // I: =H{r}=G{r}
            c += `<c r="I${r}" s="211"><f>H${r}=G${r}</f></c>`
            return `<row r="${r}" spans="1:15" x14ac:dyDescent="0.3">${c}</row>`
          }).join('')
          let auxAlimXml = strFromU8(files[auxAlimPath])
          // Inyectar/reemplazar F2 con fórmula =Parametros!D13
          const f2New = '<c r="F2"><f>Parametros!D13</f></c>'
          if (/<c r="F2"[^>]*>[\s\S]*?<\/c>|<c r="F2"[^/]*\/>/.test(auxAlimXml)) {
            auxAlimXml = auxAlimXml.replace(/<c r="F2"[^>]*>[\s\S]*?<\/c>|<c r="F2"[^/]*\/>/, f2New)
          } else {
            auxAlimXml = auxAlimXml.replace(/(<row r="2"[^>]*>)([\s\S]*?)(<\/row>)/, (_, open, cells, close) => open + cells + f2New + close)
          }
          files[auxAlimPath] = strToU8(rebuildSheetData(auxAlimXml, auxAlimRowsXml, 4))
        }
      }

      // ── Procesar Validacion novedades ────────────────────────────────────────
      if (fileNomina && nominaRows.length > 0) {
        const vnPath = resolveSheetZipPath(workbookXml, relsXml, 'Validacion novedades')
        if (files[vnPath]) {
          // Ordenar por DESCRIPCION CONCEPTO (col 4) A→Z
          const sortedNomina = [...nominaRows].sort((a, b) => {
            const da = String(getCellValue(a.getCell(4)) ?? '').trim()
            const db = String(getCellValue(b.getCell(4)) ?? '').trim()
            return da.localeCompare(db, undefined, { sensitivity: 'base' })
          })
          const lastDataRow = 3 + sortedNomina.length
          const vnRowsXml = sortedNomina.map((srcRow, idx) => {
            const r = idx + 4 // datos desde fila 4 (encabezados en fila 3)
            const descConceptoVal = getCellValue(srcRow.getCell(4))
            const cantidadVal = String(descConceptoVal ?? '').trim() === 'SIP Payment'
              ? 0
              : getCellValue(srcRow.getCell(5))
            return buildValNovedadesRow(
              r,
              getCellValue(srcRow.getCell(1)), // CODIGO EMPLEADO → col B
              getCellValue(srcRow.getCell(2)), // NOMBRES EMPLEADO → col C
              getCellValue(srcRow.getCell(3)), // CONCEPTO → col D
              descConceptoVal,                 // NOMBRE CONCEPTO → col E
              cantidadVal,                     // CANTIDAD → col F
              getCellValue(srcRow.getCell(6)), // DEVENGOS → col G
              getCellValue(srcRow.getCell(7)), // DEDUCCIONES → col H
              strIdxMap
            )
          }).join('')
          let vnXml = rebuildSheetData(strFromU8(files[vnPath]), vnRowsXml, 3)
          // Actualizar rangos del autoFilter y sortState al número real de filas
          vnXml = vnXml
            .replace(/(<autoFilter\b[^>]*\bref=")[^"]*(")/,   `$1A3:N${lastDataRow}$2`)
            .replace(/(<sortState\b[^>]*\bref=")[^"]*(")/,    `$1A4:N${lastDataRow}$2`)
            .replace(/(<sortCondition\b[^>]*\bref=")[^"]*(")/,`$1E3:E${lastDataRow}$2`)
          files[vnPath] = strToU8(vnXml)
        }
      }

      // ── Procesar Novedades del cliente ("All Entitlements" → hoja "Novedades") ─
      if (fileNovedades) {
        const novClienteBuf = await fileNovedades.arrayBuffer()
        const novWb = new ExcelJS.Workbook()
        await novWb.xlsx.load(novClienteBuf)
        const allEnt = novWb.getWorksheet('All Entitlements')
        if (!allEnt) throw new Error('No se encontró la hoja "All Entitlements" en el archivo de Novedades.')

        // Recoger filas desde fila 2 (fila 1 = encabezados)
        const novEntRows = []
        allEnt.eachRow({ includeEmpty: false }, (row, rowNum) => {
          if (rowNum < 2) return
          const colA = getCellValue(row.getCell(1))  // Employee No
          const colD = getCellValue(row.getCell(4))  // Element Type
          const colK = getCellValue(row.getCell(11)) // col K (fallback para F)
          const colN = getCellValue(row.getCell(14)) // No. Of Units
          const colS = getCellValue(row.getCell(19)) // Task ID
          if (colA != null && String(colA).trim()) novEntRows.push({ colA, colD, colK, colN, colS })
        })

        if (novEntRows.length > 0) {
          const novPath = resolveSheetZipPath(workbookXml, relsXml, 'Novedades')
          if (novPath && files[novPath]) {
            // Construir XML de filas: A=empleado, B/C/E/G/H/I/J/K=fórmulas, D=código concepto, F=cantidad, L=Task ID
            const novRowsXml = novEntRows.map(({ colA, colD, colK, colN, colS }, idx) => {
              const r = idx + 2 // datos desde fila 2
              const empStr = String(colA).trim()
              const empIsNum = /^\d+$/.test(empStr)
              let c = ''
              // A: código empleado
              if (empIsNum) {
                c += `<c r="A${r}" s="114"><v>${empStr}</v></c>`
              } else {
                const si = strIdxMap.get(empStr)
                if (si !== undefined) c += `<c r="A${r}" s="114" t="s"><v>${si}</v></c>`
              }
              // B, C: VLOOKUP por código empleado
              c += `<c r="B${r}" t="str"><f>VLOOKUP(A${r},'Maestro Personal'!$B:$F,4,0)</f></c>`
              c += `<c r="C${r}" t="str"><f>VLOOKUP(A${r},'Maestro Personal'!$B:$F,5,0)</f></c>`
              // D: código concepto (mapeado desde Element Type)
              const elemType = colD != null ? String(colD).trim() : ''
              const conceptoCod = ELEMENT_TYPE_CODE_MAP.get(elemType)
              if (conceptoCod) {
                const si = strIdxMap.get(conceptoCod)
                if (si !== undefined) {
                  c += `<c r="D${r}" s="84" t="s"><v>${si}</v></c>`
                } else {
                  c += `<c r="D${r}" s="84" t="inlineStr"><is><t>${escapeXml(conceptoCod)}</t></is></c>`
                }
              }
              // E: VLOOKUP por concepto
              c += `<c r="E${r}" t="str"><f>VLOOKUP(D${r},Conceptos!$A:$B,2,0)</f></c>`
              // F: cantidad — col N si tiene valor; si no, col K si tiene valor (incluso 0); si no, vacío
              const nFromN = typeof colN === 'number' ? colN : (colN != null && String(colN).trim() !== '' ? parseFloat(colN) : NaN)
              const nFromK = typeof colK === 'number' ? colK : (colK != null && String(colK).trim() !== '' ? parseFloat(colK) : NaN)
              const rawF = !isNaN(nFromN) ? nFromN : (!isNaN(nFromK) ? nFromK : NaN)
              if (!isNaN(rawF)) {
                const finalF = conceptoCod === '100015' ? rawF * 80000 : rawF
                c += `<c r="F${r}"><v>${finalF}</v></c>`
              }
              // G: = F (copia cantidad)
              c += `<c r="G${r}" s="45"><f>F${r}</f></c>`
              // H: concatenar empleado + concepto
              c += `<c r="H${r}" t="str"><f>+CONCATENATE(A${r},D${r})</f></c>`
              // I: SUMIF
              c += `<c r="I${r}" s="45"><f>+SUMIF(H:H,CONCATENATE(A${r},D${r}),G:G)</f></c>`
              // J: IFERROR VLOOKUP validacion novedades
              c += `<c r="J${r}"><f>+IFERROR(VLOOKUP(H${r},'Validacion novedades'!$A:$J,10,0),0)</f></c>`
              // K: diferencia I - J
              c += `<c r="K${r}" s="47"><f>+I${r}-J${r}</f></c>`
              // L: Task ID (columna S del fuente)
              const colSStr = colS != null ? String(colS).trim() : ''
              if (colSStr) c += `<c r="L${r}" t="inlineStr"><is><t>${escapeXml(colSStr)}</t></is></c>`
              return c ? `<row r="${r}" spans="1:13" x14ac:dyDescent="0.3">${c}</row>` : ''
            }).filter(Boolean).join('')

            const novLastRow = 1 + novEntRows.length
            let novXml = strFromU8(files[novPath])
            novXml = novXml
              .replace(/<dimension ref="[^"]*"\/>/, `<dimension ref="A1:M${novLastRow}"/>`)
            novXml = rebuildSheetData(novXml, novRowsXml, 1)
            files[novPath] = strToU8(novXml)
          }
        }
      }

      // ── Actualizar col A de "Seguridad Social", "Devengos" y "Retencion" ──────
      if (fileMaestro && maestroRows.length > 0) {
        const empCodes = maestroRows
          .map(row => String(getCellValue(row.getCell(2)) ?? '').trim())
          .filter(Boolean)

        const ssPath = resolveSheetZipPath(workbookXml, relsXml, 'Seguridad Social')
        if (files[ssPath]) {
          files[ssPath] = strToU8(patchSegSocialColG(patchSegSocialColA(strFromU8(files[ssPath]), empCodes), empCodes.length))
        }

        const devPath = resolveSheetZipPath(workbookXml, relsXml, 'Devengos')
        if (files[devPath]) {
          // Devengos: datos desde fila 4 hasta 362, col A s="81"; col Z con SUMAR.SI
          files[devPath] = strToU8(patchDevengosColZ(patchSheetColA(strFromU8(files[devPath]), empCodes, 4, 362, '81'), empCodes.length))
        }

        const retPath = resolveSheetZipPath(workbookXml, relsXml, 'Retencion')
        if (files[retPath]) {
          // Retencion: datos desde fila 2 hasta 360, col A s="5"
          files[retPath] = strToU8(patchSheetColA(strFromU8(files[retPath]), empCodes, 2, 360, '5'))
        }

        const guPath = resolveSheetZipPath(workbookXml, relsXml, 'Gross up')
        if (files[guPath]) {
          // Gross up: datos desde fila 2, col A s="36"
          // Calcular la última fila real del template para que patchSheetColA
          // sepa cuántas filas ya existen y cuántas nuevas debe crear
          const guXml = strFromU8(files[guPath])
          const guRowNums = [...guXml.matchAll(/<row r="(\d+)"/g)].map(m => parseInt(m[1]))
          const guLastTplRow = guRowNums.length > 1 ? Math.max(...guRowNums) : 65
          files[guPath] = strToU8(patchSheetColA(guXml, empCodes, 2, guLastTplRow, '36'))
        }
      }

      // ── Aplicar parámetros configurados por el usuario ───────────────────────
      const activeParams = params ?? PARAMS_DEFAULTS
      const paramPath = resolveSheetZipPath(workbookXml, relsXml, 'Parametros')
      if (paramPath && files[paramPath]) {
        files[paramPath] = strToU8(patchParametrosSheet(strFromU8(files[paramPath]), activeParams))
      }

      // ── Parchar hoja "Conceptos": col C = "Incapacidad" para códigos 001150/001151/001162 ──
      const conceptosPath = resolveSheetZipPath(workbookXml, relsXml, 'Conceptos')
      if (conceptosPath && files[conceptosPath]) {
        const ssArr = buildSsReverseArray(strFromU8(files['xl/sharedStrings.xml']))
        files[conceptosPath] = strToU8(patchConceptosSheet(strFromU8(files[conceptosPath]), ssArr))
      }

      // ── Retención adicional ──────────────────────────────────────────────────
      if (activeParams.reteAdic && activeParams.reteAdic.length > 0) {
        const reteStrings = activeParams.reteAdic.map(rec => String(rec.nombre)).filter(Boolean)
        const { strIdxMap, newSsXml } = buildSharedStringIndex(strFromU8(files['xl/sharedStrings.xml']), reteStrings)
        files['xl/sharedStrings.xml'] = strToU8(newSsXml)
        const reteAdicPath = resolveSheetZipPath(workbookXml, relsXml, 'rete adic')
        if (reteAdicPath && files[reteAdicPath]) {
          files[reteAdicPath] = strToU8(buildReteAdicSheet(strFromU8(files[reteAdicPath]), activeParams.reteAdic, strIdxMap))
        }
      }

      // ── Generar archivos ─────────────────────────────────────────────────────
      delete files['xl/calcChain.xml']
      // Forzar recálculo al abrir en ambos archivos (Seguridad Social depende de Detallado Mes)
      files['xl/workbook.xml'] = strToU8(addFullCalcOnLoad(workbookXml))
      const fullBytes = zipSync(files)
      const updatedWbXml = strFromU8(files['xl/workbook.xml'])
      const [segSocialBytes, retencionBytes] = await Promise.all([
        extractSegSocial(files, updatedWbXml, relsXml),
        extractSheetStandalone(files, updatedWbXml, relsXml, 'Retencion'),
      ])

      const parts = []
      if (file) parts.push(`${dataRows.length} registros IBC`)
      if (fileNomina) parts.push(`${nominaRows.length} registros nómina`)
      if (fileMaestro) parts.push(`${maestroRows.length} empleados maestro`)
      setModal({
        type: 'success',
        message: `Se procesaron ${parts.join(' y ')} correctamente.`,
        blobs: { full: fullBytes, segSocial: segSocialBytes, retencion: retencionBytes },
      })
    } catch (err) {
      console.error(err)
      setModal({ type: 'error', message: err.message || 'Error al procesar el archivo.' })
    } finally {
      setIsProcessing(false)
    }
  }

  return (
    <div className="app">
      <header className="header">
        <div className="container">
          <div className="header-content">
            <div className="logo-container">
              <div className="logo">
                <img src="/Logo syp.png" alt="Solutions & Payroll Logo" width="60" height="60" />
              </div>
              <div className="header-text">
                <h1>Solutions & Payroll</h1>
                <p className="subtitle">Validación de Nómina</p>
              </div>
            </div>
            <div className="welcome-box">
              <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                <path d="M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2"/>
                <circle cx="12" cy="7" r="4"/>
              </svg>
              <span>Bienvenido, Usuario</span>
            </div>
          </div>
        </div>
      </header>

      <main className="main-content">
        <div className="container">

          <div className="tab-nav">
            <button
              className={`tab-btn ${activeTab === 'validacion' ? 'active' : ''}`}
              onClick={() => setActiveTab('validacion')}
            >
              <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
                <polyline points="14 2 14 8 20 8"/>
              </svg>
              Validación
            </button>
            <button
              className={`tab-btn ${activeTab === 'parametros' ? 'active' : ''}`}
              onClick={() => setActiveTab('parametros')}
            >
              <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                <circle cx="12" cy="12" r="3"/>
                <path d="M19.07 4.93a10 10 0 0 1 0 14.14M4.93 4.93a10 10 0 0 0 0 14.14"/>
              </svg>
              Parámetros
            </button>
          </div>

          {activeTab === 'validacion' && <>

          <div className="help-section">
            <button
              className="help-toggle"
              onClick={() => setIsHelpExpanded(!isHelpExpanded)}
              aria-expanded={isHelpExpanded}
            >
              <div className="help-toggle-header">
                <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                  <circle cx="12" cy="12" r="10"/>
                  <line x1="12" y1="16" x2="12" y2="12"/>
                  <line x1="12" y1="8" x2="12.01" y2="8"/>
                </svg>
                <span>¿Cómo usar esta aplicación?</span>
              </div>
              <svg
                className={`chevron ${isHelpExpanded ? 'expanded' : ''}`}
                width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"
              >
                <polyline points="6 9 12 15 18 9"/>
              </svg>
            </button>
            <div className={`help-content ${isHelpExpanded ? 'expanded' : ''}`}>
              <ol className="help-list">
                <li>
                  <span className="step-number">1</span>
                  <div>
                    <strong>Plantilla base (opcional)</strong>
                    <p>Si tienes una plantilla personalizada, súbela aquí. Si no, la app usa la plantilla interna por defecto.</p>
                  </div>
                </li>
                <li>
                  <span className="step-number">2</span>
                  <div>
                    <strong>IBC Mes Anterior</strong>
                    <p>El Excel de IBC del mes anterior. Debe contener la hoja <em>Sheet1</em>. La app mapea más de 80 columnas y construye el pivote de bases IBC en la hoja <em>Planilla Mes anterior</em>.</p>
                  </div>
                </li>
                <li>
                  <span className="step-number">3</span>
                  <div>
                    <strong>Archivo Nómina</strong>
                    <p>El Excel de nómina detallada del mes actual. Se usa para llenar <em>Detallado Mes</em>, <em>rev incapacidades</em> (conceptos 001150 / 001151 / 001177 / 001178) y <em>Validacion novedades</em>.</p>
                  </div>
                </li>
                <li>
                  <span className="step-number">4</span>
                  <div>
                    <strong>Nómina mes anterior (opcional)</strong>
                    <p>La nómina del mes pasado con la misma estructura. Se usa para determinar qué concepto de auxilio de alimentación (111500 ó 111501) tenía cada empleado el mes anterior y volcar ese dato en la columna H de <em>validacion aux alimentacion (2)</em>. Si no se sube, esa columna queda con una fórmula BUSCARV.</p>
                  </div>
                </li>
                <li>
                  <span className="step-number">5</span>
                  <div>
                    <strong>Maestro Personal</strong>
                    <p>El Excel de Maestro Personal actualizado. Se toman los empleados con estado <em>Activos</em>, <em>Ausencia</em> o <em>Vacaciones</em> (desde la fila 6). Sus datos llenan la hoja <em>Maestro Personal</em>, actualizan la columna A de <em>Seguridad Social</em> y construyen la hoja <em>validacion aux alimentacion (2)</em>.</p>
                  </div>
                </li>
                <li>
                  <span className="step-number">6</span>
                  <div>
                    <strong>Novedades del cliente (opcional)</strong>
                    <p>El Excel de novedades exportado desde el sistema del cliente (<em>All Entitlements</em>). Se mapea a la hoja <em>Novedades</em> y es la fuente del BUSCARV en la columna K de <em>Validacion novedades</em>.</p>
                  </div>
                </li>
                <li>
                  <span className="step-number">7</span>
                  <div>
                    <strong>Generar y descargar</strong>
                    <p>Haz clic en <strong>Generar</strong>. Se descargará el <strong>Libro completo</strong> — la plantilla entera con todas las hojas y los datos del mes cargados, lista para abrir en Excel.</p>
                  </div>
                </li>
              </ol>
            </div>
          </div>

          <div className="card">
            <div className="card-header">
              <h2>Generar Planilla de Validación</h2>
              <p className="description">
                Sube el archivo Excel de IBC del mes anterior para generar la planilla con los datos mapeados.
              </p>
            </div>

            <div className="card-body">
              <div className="form-section">

                <div className="form-group form-group--plantilla">
                  <label className="label label--plantilla">
                    <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                      <rect x="3" y="3" width="18" height="18" rx="2"/><path d="M3 9h18M9 21V9"/>
                    </svg>
                    Plantilla base (.xlsx)
                    <span className="label-badge">Requerido</span>
                  </label>
                  <p className="plantilla-hint">Sube la plantilla del mes anterior para que los datos se acumulen correctamente en la hoja BVAC. Si no subes ninguna, se usará la plantilla por defecto.</p>
                  <input
                    ref={filePlantillaRef}
                    type="file"
                    accept=".xlsx,.xls"
                    className="file-input"
                    id="file-upload-plantilla"
                    onChange={(e) => handleFilePlantilla(e.target.files[0])}
                  />
                  {!filePlantilla ? (
                    <label
                      htmlFor="file-upload-plantilla"
                      className={`drop-zone drop-zone--plantilla ${isDraggingPlantilla ? 'drag-active' : ''}`}
                      onDragOver={(e) => { e.preventDefault(); setIsDraggingPlantilla(true) }}
                      onDragLeave={() => setIsDraggingPlantilla(false)}
                      onDrop={handleDropPlantilla}
                    >
                      <div className="drop-zone-content">
                        <svg width="48" height="48" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5">
                          <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
                          <polyline points="17 8 12 3 7 8"/>
                          <line x1="12" y1="3" x2="12" y2="15"/>
                        </svg>
                        <div className="drop-zone-text">
                          <span className="drop-zone-title">Arrastra la plantilla aquí</span>
                          <span className="drop-zone-subtitle">o haz clic para seleccionarla</span>
                        </div>
                        <span className="drop-zone-hint">Formatos aceptados: .xlsx, .xls</span>
                      </div>
                    </label>
                  ) : (
                    <div className="drop-zone drop-zone--plantilla has-file">
                      <div className="file-preview">
                        <div className="file-icon file-icon--plantilla">
                          <svg width="32" height="32" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5">
                            <rect x="3" y="3" width="18" height="18" rx="2"/><path d="M3 9h18M9 21V9"/>
                          </svg>
                        </div>
                        <div className="file-details">
                          <span className="file-name">{filePlantilla.name}</span>
                          <span className="file-size">{formatBytes(filePlantilla.size)}</span>
                        </div>
                        <button className="btn-remove" onClick={removeFilePlantilla} type="button" aria-label="Eliminar plantilla">
                          <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                            <line x1="18" y1="6" x2="6" y2="18"/>
                            <line x1="6" y1="6" x2="18" y2="18"/>
                          </svg>
                        </button>
                      </div>
                    </div>
                  )}
                </div>

                <div className="form-group">
                  <label className="label">
                    <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                      <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
                      <polyline points="14 2 14 8 20 8"/>
                    </svg>
                    Archivo IBC Mes Anterior (.xlsx)
                  </label>
                  <input
                    ref={fileInputRef}
                    type="file"
                    accept=".xlsx,.xls"
                    className="file-input"
                    id="file-upload"
                    onChange={(e) => handleFile(e.target.files[0])}
                  />
                  {!file ? (
                    <label
                      htmlFor="file-upload"
                      className={`drop-zone ${isDragging ? 'drag-active' : ''}`}
                      onDragOver={(e) => { e.preventDefault(); setIsDragging(true) }}
                      onDragLeave={() => setIsDragging(false)}
                      onDrop={handleDrop}
                    >
                      <div className="drop-zone-content">
                        <svg width="48" height="48" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5">
                          <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
                          <polyline points="17 8 12 3 7 8"/>
                          <line x1="12" y1="3" x2="12" y2="15"/>
                        </svg>
                        <div className="drop-zone-text">
                          <span className="drop-zone-title">Arrastra tu archivo aquí</span>
                          <span className="drop-zone-subtitle">o haz clic para seleccionarlo</span>
                        </div>
                        <span className="drop-zone-hint">Formatos aceptados: .xlsx, .xls</span>
                      </div>
                    </label>
                  ) : (
                    <div className="drop-zone has-file">
                      <div className="file-preview">
                        <div className="file-icon">
                          <svg width="32" height="32" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5">
                            <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
                            <polyline points="14 2 14 8 20 8"/>
                          </svg>
                        </div>
                        <div className="file-details">
                          <span className="file-name">{file.name}</span>
                          <span className="file-size">{formatBytes(file.size)}</span>
                        </div>
                        <button className="btn-remove" onClick={removeFile} type="button" aria-label="Eliminar archivo">
                          <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                            <line x1="18" y1="6" x2="6" y2="18"/>
                            <line x1="6" y1="6" x2="18" y2="18"/>
                          </svg>
                        </button>
                      </div>
                    </div>
                  )}
                </div>

                <div className="form-group">
                  <label className="label">
                    <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                      <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
                      <polyline points="14 2 14 8 20 8"/>
                    </svg>
                    Archivo Nómina (.xlsx)
                  </label>
                  <input
                    ref={fileNominaRef}
                    type="file"
                    accept=".xlsx,.xls"
                    className="file-input"
                    id="file-upload-nomina"
                    onChange={(e) => handleFileNomina(e.target.files[0])}
                  />
                  {!fileNomina ? (
                    <label
                      htmlFor="file-upload-nomina"
                      className={`drop-zone ${isDraggingNomina ? 'drag-active' : ''}`}
                      onDragOver={(e) => { e.preventDefault(); setIsDraggingNomina(true) }}
                      onDragLeave={() => setIsDraggingNomina(false)}
                      onDrop={handleDropNomina}
                    >
                      <div className="drop-zone-content">
                        <svg width="48" height="48" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5">
                          <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
                          <polyline points="17 8 12 3 7 8"/>
                          <line x1="12" y1="3" x2="12" y2="15"/>
                        </svg>
                        <div className="drop-zone-text">
                          <span className="drop-zone-title">Arrastra tu archivo aquí</span>
                          <span className="drop-zone-subtitle">o haz clic para seleccionarlo</span>
                        </div>
                        <span className="drop-zone-hint">Formatos aceptados: .xlsx, .xls</span>
                      </div>
                    </label>
                  ) : (
                    <div className="drop-zone has-file">
                      <div className="file-preview">
                        <div className="file-icon">
                          <svg width="32" height="32" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5">
                            <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
                            <polyline points="14 2 14 8 20 8"/>
                          </svg>
                        </div>
                        <div className="file-details">
                          <span className="file-name">{fileNomina.name}</span>
                          <span className="file-size">{formatBytes(fileNomina.size)}</span>
                        </div>
                        <button className="btn-remove" onClick={removeFileNomina} type="button" aria-label="Eliminar archivo">
                          <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                            <line x1="18" y1="6" x2="6" y2="18"/>
                            <line x1="6" y1="6" x2="18" y2="18"/>
                          </svg>
                        </button>
                      </div>
                    </div>
                  )}
                </div>

                {/* Nómina mes anterior */}
                <div className="form-group">
                  <label className="label">
                    <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                      <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
                      <polyline points="14 2 14 8 20 8"/>
                    </svg>
                    Nómina mes anterior (.xlsx)
                    <span className="drop-zone-hint" style={{display:'block', fontWeight:'normal', fontSize:'0.8rem', marginTop:'4px'}}>Para determinar el concepto de auxilio de alimentación usado el mes pasado por cada persona.</span>
                  </label>
                  <input
                    ref={fileNominaAnteriorRef}
                    type="file"
                    accept=".xlsx,.xls"
                    className="file-input"
                    id="file-upload-nomina-anterior"
                    onChange={(e) => handleFileNominaAnterior(e.target.files[0])}
                  />
                  {!fileNominaAnterior ? (
                    <label
                      htmlFor="file-upload-nomina-anterior"
                      className={`drop-zone ${isDraggingNominaAnterior ? 'drag-active' : ''}`}
                      onDragOver={(e) => { e.preventDefault(); setIsDraggingNominaAnterior(true) }}
                      onDragLeave={() => setIsDraggingNominaAnterior(false)}
                      onDrop={handleDropNominaAnterior}
                    >
                      <div className="drop-zone-content">
                        <svg width="48" height="48" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5">
                          <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
                          <polyline points="17 8 12 3 7 8"/>
                          <line x1="12" y1="3" x2="12" y2="15"/>
                        </svg>
                        <div className="drop-zone-text">
                          <span className="drop-zone-title">Arrastra tu archivo aquí</span>
                          <span className="drop-zone-subtitle">o haz clic para seleccionarlo</span>
                        </div>
                        <span className="drop-zone-hint">Misma estructura que la nómina actual (opcional)</span>
                      </div>
                    </label>
                  ) : (
                    <div className="drop-zone has-file">
                      <div className="file-preview">
                        <div className="file-icon">
                          <svg width="32" height="32" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5">
                            <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
                            <polyline points="14 2 14 8 20 8"/>
                          </svg>
                        </div>
                        <div className="file-details">
                          <span className="file-name">{fileNominaAnterior.name}</span>
                          <span className="file-size">{formatBytes(fileNominaAnterior.size)}</span>
                        </div>
                        <button className="btn-remove" onClick={removeFileNominaAnterior} type="button" aria-label="Eliminar archivo">
                          <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                            <line x1="18" y1="6" x2="6" y2="18"/>
                            <line x1="6" y1="6" x2="18" y2="18"/>
                          </svg>
                        </button>
                      </div>
                    </div>
                  )}
                </div>

                <div className="form-group">
                  <label className="label">
                    <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                      <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
                      <polyline points="14 2 14 8 20 8"/>
                    </svg>
                    Maestro Personal (.xlsx)
                  </label>
                  <input
                    ref={fileMaestroRef}
                    type="file"
                    accept=".xlsx,.xls"
                    className="file-input"
                    id="file-upload-maestro"
                    onChange={(e) => handleFileMaestro(e.target.files[0])}
                  />
                  {!fileMaestro ? (
                    <label
                      htmlFor="file-upload-maestro"
                      className={`drop-zone ${isDraggingMaestro ? 'drag-active' : ''}`}
                      onDragOver={(e) => { e.preventDefault(); setIsDraggingMaestro(true) }}
                      onDragLeave={() => setIsDraggingMaestro(false)}
                      onDrop={handleDropMaestro}
                    >
                      <div className="drop-zone-content">
                        <svg width="48" height="48" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5">
                          <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
                          <polyline points="17 8 12 3 7 8"/>
                          <line x1="12" y1="3" x2="12" y2="15"/>
                        </svg>
                        <div className="drop-zone-text">
                          <span className="drop-zone-title">Arrastra tu archivo aquí</span>
                          <span className="drop-zone-subtitle">o haz clic para seleccionarlo</span>
                        </div>
                        <span className="drop-zone-hint">Formatos aceptados: .xlsx, .xls</span>
                      </div>
                    </label>
                  ) : (
                    <div className="drop-zone has-file">
                      <div className="file-preview">
                        <div className="file-icon">
                          <svg width="32" height="32" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5">
                            <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
                            <polyline points="14 2 14 8 20 8"/>
                          </svg>
                        </div>
                        <div className="file-details">
                          <span className="file-name">{fileMaestro.name}</span>
                          <span className="file-size">{formatBytes(fileMaestro.size)}</span>
                        </div>
                        <button className="btn-remove" onClick={removeFileMaestro} type="button" aria-label="Eliminar archivo">
                          <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                            <line x1="18" y1="6" x2="6" y2="18"/>
                            <line x1="6" y1="6" x2="18" y2="18"/>
                          </svg>
                        </button>
                      </div>
                    </div>
                  )}
                </div>

                <div className="form-group">
                  <label className="label">
                    <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                      <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
                      <polyline points="14 2 14 8 20 8"/>
                    </svg>
                    Novedades del cliente (.xlsx)
                  </label>
                  <input
                    ref={fileNovedadesRef}
                    type="file"
                    accept=".xlsx,.xls"
                    className="file-input"
                    id="file-upload-novedades"
                    onChange={(e) => handleFileNovedades(e.target.files[0])}
                  />
                  {!fileNovedades ? (
                    <label
                      htmlFor="file-upload-novedades"
                      className={`drop-zone ${isDraggingNovedades ? 'drag-active' : ''}`}
                      onDragOver={(e) => { e.preventDefault(); setIsDraggingNovedades(true) }}
                      onDragLeave={() => setIsDraggingNovedades(false)}
                      onDrop={handleDropNovedades}
                    >
                      <div className="drop-zone-content">
                        <svg width="48" height="48" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5">
                          <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
                          <polyline points="17 8 12 3 7 8"/>
                          <line x1="12" y1="3" x2="12" y2="15"/>
                        </svg>
                        <div className="drop-zone-text">
                          <span className="drop-zone-title">Arrastra tu archivo aquí</span>
                          <span className="drop-zone-subtitle">o haz clic para seleccionarlo</span>
                        </div>
                        <span className="drop-zone-hint">Hoja requerida: "All Entitlements"</span>
                      </div>
                    </label>
                  ) : (
                    <div className="drop-zone has-file">
                      <div className="file-preview">
                        <div className="file-icon">
                          <svg width="32" height="32" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5">
                            <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
                            <polyline points="14 2 14 8 20 8"/>
                          </svg>
                        </div>
                        <div className="file-details">
                          <span className="file-name">{fileNovedades.name}</span>
                          <span className="file-size">{formatBytes(fileNovedades.size)}</span>
                        </div>
                        <button className="btn-remove" onClick={removeFileNovedades} type="button" aria-label="Eliminar archivo">
                          <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                            <line x1="18" y1="6" x2="6" y2="18"/>
                            <line x1="6" y1="6" x2="18" y2="18"/>
                          </svg>
                        </button>
                      </div>
                    </div>
                  )}
                </div>

                <button
                  className="btn-primary"
                  onClick={processFile}
                  disabled={(!file && !fileNomina && !fileMaestro && !filePlantilla && !fileNovedades && !fileNominaAnterior) || isProcessing}
                >
                  {isProcessing ? (
                    <>
                      <svg className="spinner" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                        <path d="M21 12a9 9 0 1 1-6.219-8.56"/>
                      </svg>
                      Procesando...
                    </>
                  ) : (
                    <>
                      <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                        <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
                        <polyline points="7 10 12 15 17 10"/>
                        <line x1="12" y1="15" x2="12" y2="3"/>
                      </svg>
                      Generar
                    </>
                  )}
                </button>

                {modal && (
                  <div className={`result-section ${modal.type}`}>
                    {modal.type === 'success' ? (
                      <>
                        <div className="result-header">
                          <div className="result-icon">
                            <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5">
                              <polyline points="20 6 9 17 4 12"/>
                            </svg>
                          </div>
                          <div>
                            <h3 className="result-title">¡Listo!</h3>
                            <p className="result-message">{modal.message}</p>
                          </div>
                        </div>
                        {modal.blobs && (
                          <div className="result-downloads">
                            {/* Seguridad Social y Retención ocultos temporalmente */}
                            <button
                              className="result-download-btn"
                              onClick={() => downloadBlob(modal.blobs.full, 'Validacion_Nomina.xlsx')}
                            >
                              <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                                <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
                                <polyline points="7 10 12 15 17 10"/>
                                <line x1="12" y1="15" x2="12" y2="3"/>
                              </svg>
                              Libro completo
                            </button>
                          </div>
                        )}
                      </>
                    ) : (
                      <div className="result-error-row">
                        <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                          <circle cx="12" cy="12" r="10"/>
                          <line x1="12" y1="8" x2="12" y2="12"/>
                          <line x1="12" y1="16" x2="12.01" y2="16"/>
                        </svg>
                        <span>{modal.message}</span>
                      </div>
                    )}
                  </div>
                )}

              </div>
            </div>
          </div>
          </>}

          {activeTab === 'parametros' && (
            <div className="card">
              <div className="card-header">
                <h2>Parámetros de Nómina</h2>
                <p className="description">Configura los valores que se usarán en el Excel generado. Los cambios se aplican automáticamente al generar la planilla.</p>
              </div>
              <div className="card-body">
                {params === null ? (
                  <p style={{color: 'var(--text-secondary)', padding: '1rem 0'}}>Cargando parámetros de la plantilla...</p>
                ) : (
                  <div className="params-content">
                    <div className="params-group">
                      <h3 className="params-group-title">Valor UVT</h3>
                      <div className="params-uvt-row">
                        <label className="label">VALOR UVT</label>
                        <input
                          type="number"
                          className="params-input params-input-uvt"
                          value={params.uvt}
                          onChange={e => setParams(p => ({...p, uvt: Number(e.target.value)}))}
                          min="0"
                        />
                      </div>
                    </div>

                    <div className="params-group">
                      <h3 className="params-group-title">Tabla de Retención</h3>
                      <table className="params-table">
                        <thead>
                          <tr>
                            <th>Valor 1</th>
                            <th>Valor 2</th>
                            <th>Valor 3</th>
                          </tr>
                        </thead>
                        <tbody>
                          {params.ret.map((row, i) => (
                            <tr key={i}>
                              <td><input type="number" className="params-input" value={row.desde}
                                onChange={e => setParams(p => ({...p, ret: p.ret.map((r,j) => j===i ? {...r, desde: Number(e.target.value)} : r)}))}
                                min="0" /></td>
                              <td><input type="number" className="params-input" value={+(row.tarifa*100).toFixed(4)}
                                onChange={e => setParams(p => ({...p, ret: p.ret.map((r,j) => j===i ? {...r, tarifa: Number(e.target.value)/100} : r)}))}
                                min="0" max="100" step="0.01" /></td>
                              <td><input type="number" className="params-input" value={row.descuento}
                                onChange={e => setParams(p => ({...p, ret: p.ret.map((r,j) => j===i ? {...r, descuento: Number(e.target.value)} : r)}))}
                                min="0" /></td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>

                    {/* ── Retención Adicional ─────────────────────────────── */}
                    <div className="params-group">
                      <div className="params-group-header">
                        <h3 className="params-group-title">Retención Adicional</h3>
                        <button
                          className="rete-add-btn"
                          onClick={() => setParams(p => ({
                            ...p,
                            reteAdic: [...(p.reteAdic ?? []), { empleado: '', nombre: '', retAdicional: 0 }]
                          }))}
                        >
                          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5">
                            <line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/>
                          </svg>
                          Agregar registro
                        </button>
                      </div>
                      {params.reteAdic && params.reteAdic.length > 0 ? (
                        <table className="params-table rete-table">
                          <thead>
                            <tr>
                              <th>Empleado</th>
                              <th>Nombre completo</th>
                              <th>RET ADICIONAL</th>
                              <th></th>
                            </tr>
                          </thead>
                          <tbody>
                            {params.reteAdic.map((row, i) => (
                              <tr key={i}>
                                <td>
                                  <input
                                    type="text"
                                    className="params-input"
                                    value={row.empleado}
                                    onChange={e => setParams(p => ({
                                      ...p,
                                      reteAdic: p.reteAdic.map((r, j) => j === i ? {...r, empleado: e.target.value} : r)
                                    }))}
                                    placeholder="Cédula / Código"
                                  />
                                </td>
                                <td>
                                  <input
                                    type="text"
                                    className="params-input"
                                    value={row.nombre}
                                    onChange={e => setParams(p => ({
                                      ...p,
                                      reteAdic: p.reteAdic.map((r, j) => j === i ? {...r, nombre: e.target.value} : r)
                                    }))}
                                    placeholder="Nombre completo"
                                  />
                                </td>
                                <td>
                                  <input
                                    type="number"
                                    className="params-input"
                                    value={row.retAdicional}
                                    onChange={e => setParams(p => ({
                                      ...p,
                                      reteAdic: p.reteAdic.map((r, j) => j === i ? {...r, retAdicional: Number(e.target.value)} : r)
                                    }))}
                                    min="0"
                                  />
                                </td>
                                <td>
                                  <button
                                    className="rete-del-btn"
                                    onClick={() => setParams(p => ({
                                      ...p,
                                      reteAdic: p.reteAdic.filter((_, j) => j !== i)
                                    }))}
                                    title="Eliminar"
                                  >
                                    <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5">
                                      <line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/>
                                    </svg>
                                  </button>
                                </td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      ) : (
                        <p className="rete-empty">Sin registros. Haz clic en "Agregar registro" para añadir uno.</p>
                      )}
                    </div>
                  </div>
                )}
              </div>
            </div>
          )}

        </div>
      </main>

      <footer className="footer">
        <div className="container">
          <p>&copy; {new Date().getFullYear()} Solutions & Payroll. Todos los derechos reservados.</p>
        </div>
      </footer>


    </div>
  )
}

export default App
