import { useState, useRef } from 'react'
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

function buildRevIncapRow(r, empleado, nombre, concepto, descConcepto, cantidad, devengos, deducciones, strIdxMap) {
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
  const cant = Number(cantidad ?? 0);    c += `<c r="F${r}"><v>${isNaN(cant) ? 0 : cant}</v></c>`
  const dev  = Number(devengos ?? 0);    c += `<c r="G${r}"><v>${isNaN(dev)  ? 0 : dev}</v></c>`
  const ded  = Number(deducciones ?? 0); c += `<c r="H${r}"><v>${isNaN(ded)  ? 0 : ded}</v></c>`
  c += `<c r="I${r}"><f>+VLOOKUP(B${r},'Maestro Personal'!$B$2:$E$312,4,0)</f></c>`
  c += `<c r="J${r}"><f>+VLOOKUP(I${r},'Planilla Mes anterior'!$CX:$CY,2,0)*100/70</f></c>`
  c += `<c r="K${r}"><f>+ROUND(J${r}/30*70%*F${r}*$C$4,0)</f></c>`
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

// ─── Actualiza col A de "Seguridad Social" con códigos de empleados ─────────
/**
 * Sobrescribe la columna A en la hoja "Seguridad Social" con los códigos de
 * empleados del Maestro Personal.
 *   · Filas 3-361: reemplaza col A (conserva todas las fórmulas/estilos intactos)
 *   · Si hay menos empleados que filas disponibles: las filas sobrantes quedan sin col A
 *   · Si hay más de 359 empleados: genera filas 362+ a partir de la plantilla de fila 3
 *     con fórmulas explícitas y número de fila actualizado
 */
function patchSegSocialColA(sheetXml, empleadoCodes) {
  const FIRST_ROW = 3
  const LAST_TPL_ROW = 361
  const total = empleadoCodes.length

  // ── Paso 1: parchear filas 3-361 (reemplazar/quitar col A) ──────────────────
  let result = sheetXml.replace(/<row r="(\d+)"([^>]*)(?<!\/)>([\s\S]*?)<\/row>/g,
    (match, rStr, attrs, inner) => {
      const r = parseInt(rStr)
      if (r < FIRST_ROW || r > LAST_TPL_ROW) return match

      // Eliminar celda A existente (tanto normal como self-closing)
      let newInner = inner
        .replace(/<c r="A\d+"([^>]*)(?<!\/)>([\s\S]*?)<\/c>/g, '')
        .replace(/<c r="A\d+"[^>]*\/>/g, '')

      const idx = r - FIRST_ROW
      if (idx < total && empleadoCodes[idx]) {
        newInner = `<c r="A${r}" s="81"><v>${empleadoCodes[idx]}</v></c>` + newInner
      }
      return `<row r="${r}"${attrs}>${newInner}</row>`
    })

  // ── Paso 2: añadir filas 362+ si hay más empleados que huecos de plantilla ─
  const TPL_SLOTS = LAST_TPL_ROW - FIRST_ROW + 1   // 359
  if (total > TPL_SLOTS) {
    const row3M = sheetXml.match(/<row r="3"([^>]*)(?<!\/)>([\s\S]*?)<\/row>/)
    if (row3M) {
      const row3Attrs = row3M[1]
      const row3Inner = row3M[2]
      const newRows = []

      for (let idx = TPL_SLOTS; idx < total; idx++) {
        const r = FIRST_ROW + idx
        const code = empleadoCodes[idx]
        if (!code) continue

        let inner = row3Inner
          // a) Renumerar atributos de ref de celda: r="X3" → r="X{r}"
          .replace(/r="([A-Z]{1,3})3"/g, (_, col) => `r="${col}${r}"`)
          // b) Quitar atributos de shared-formula en <f>, conservar texto de fórmula
          .replace(/<f ([^>]+)>([\s\S]*?)<\/f>/g, '<f>$2</f>')
          // c) Renumerar referencias de fila relativas en el texto de fórmulas
          //    Solo coincide con letra(s) seguidas de "3" no precedida de "$" ni seguida de dígito
          .replace(/(?<!\$)([A-Z]{1,3})3(?!\d)/g, (_, col) => `${col}${r}`)
          // d) Eliminar valores cacheados de celdas con fórmula (Excel recalculará)
          .replace(/(<f>[\s\S]*?<\/f>)\s*<v>[^<]*<\/v>/g, '$1')
          // e) Reemplazar col A con el nuevo código de empleado
          .replace(/<c r="A[^"]*"([^>]*)(?<!\/)>([\s\S]*?)<\/c>/,
            `<c r="A${r}" s="81"><v>${code}</v></c>`)
          .replace(/<c r="A[^"]*"[^>]*\/>/,
            `<c r="A${r}" s="81"><v>${code}</v></c>`)

        newRows.push(`<row r="${r}"${row3Attrs}>${inner}</row>`)
      }

      result = result.replace('</sheetData>', newRows.join('') + '</sheetData>')
    }
  }

  return result
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
 * Produce el XLSX standalone de "Seguridad Social":
 * - Carga todas las hojas dependientes en HyperFormula
 * - Evalúa las fórmulas de sheet19 con los datos reales (Detallado Mes, etc.)
 * - Genera un XLSX mínimo con SOLO la hoja "Seguridad Social" y valores estáticos
 */
async function extractSegSocial(files, workbookXml, relsXml) {
  // ── 1. Shared strings ──────────────────────────────────────────────────────
  const ssXml = strFromU8(files['xl/sharedStrings.xml'])
  const ssArr = buildSsReverseArray(ssXml)
  const { map: ssForwardMap } = parseSharedStrings(ssXml)

  // ── 2. Cargar HyperFormula ─────────────────────────────────────────────────
  const { HyperFormula } = await import('hyperformula')

  // ── 3. Parsear hojas estáticas (solo valores cacheados) ────────────────────
  const sheetsData = {}
  const staticSheets = [
    'Parametros', 'Maestro Personal', 'Devengos',
    'rev incapacidades', 'Pensionado', 'Validacion novedades',
    'Planilla Mes anterior', 'Conceptos',
  ]
  for (const name of staticSheets) {
    try {
      const p = resolveSheetZipPath(workbookXml, relsXml, name)
      if (files[p]) sheetsData[name] = parseSheetXmlToValues(strFromU8(files[p]), ssArr)
    } catch (_) { /* hoja no encontrada, omitir */ }
  }

  // Detallado Mes — con fórmulas (columnas J-P usan VLOOKUP contra Conceptos)
  try {
    const p = resolveSheetZipPath(workbookXml, relsXml, 'Detallado Mes')
    if (files[p]) sheetsData['Detallado Mes'] = parseSheetXmlToFormulas(strFromU8(files[p]), ssArr)
  } catch (_) {}

  // Seguridad Social — con fórmulas (es lo que evaluamos)
  const ss19Path = resolveSheetZipPath(workbookXml, relsXml, 'Seguridad Social')
  const ss19Xml = strFromU8(files[ss19Path])
  sheetsData['Seguridad Social'] = parseSheetXmlToFormulas(ss19Xml, ssArr)

  // ── 4. Instanciar HyperFormula y evaluar ───────────────────────────────────
  const hf = HyperFormula.buildFromSheets(sheetsData, { licenseKey: 'gpl-v3' })
  const ssSheetId = hf.getSheetId('Seguridad Social')
  const ssGrid = sheetsData['Seguridad Social']

  // ── 5. Extraer valores evaluados de Seguridad Social ──────────────────────
  const evalMap = new Map()
  for (let r = 0; r < ssGrid.length; r++) {
    for (let c = 0; c < (ssGrid[r] || []).length; c++) {
      try {
        const val = hf.getCellValue({ sheet: ssSheetId, row: r, col: c })
        if (val !== null && val !== undefined) {
          evalMap.set(`${colToLetter(c + 1)}${r + 1}`, val)
        }
      } catch (_) {}
    }
  }

  // ── 6. Reconstruir sheet19 con valores en lugar de fórmulas ───────────────
  // Quitar referencia a printerSettings (no existe en el XLSX standalone)
  const cleanedSheet19 = ss19Xml.replace(/\s+r:id="rId\d+"/g, (m, offset, str) => {
    // Solo quitar r:id dentro de pageSetup
    const before = str.lastIndexOf('<', offset)
    return str.slice(before, offset).includes('pageSetup') ? '' : m
  })
  const evaluatedSheet19 = replaceFormulasWithValues(cleanedSheet19, evalMap, ssForwardMap)

  // ── 7. Construir ZIP mínimo con solo Seguridad Social ─────────────────────
  const CT_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/></Types>`

  const RELS_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`

  const WB_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Seguridad Social" sheetId="1" r:id="rId1"/></sheets></workbook>`

  const WB_RELS_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/></Relationships>`

  const newFiles = {
    '[Content_Types].xml': strToU8(CT_XML),
    '_rels/.rels': strToU8(RELS_XML),
    'xl/workbook.xml': strToU8(WB_XML),
    'xl/_rels/workbook.xml.rels': strToU8(WB_RELS_XML),
    'xl/worksheets/sheet1.xml': strToU8(evaluatedSheet19),
    'xl/sharedStrings.xml': files['xl/sharedStrings.xml'],
    'xl/styles.xml': files['xl/styles.xml'],
  }
  if (files['xl/theme/theme1.xml']) newFiles['xl/theme/theme1.xml'] = files['xl/theme/theme1.xml']

  return zipSync(newFiles)
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
  const [isHelpExpanded, setIsHelpExpanded] = useState(false)
  const [isProcessing, setIsProcessing] = useState(false)
  const [modal, setModal] = useState(null)
  const fileInputRef = useRef(null)
  const fileNominaRef = useRef(null)
  const fileMaestroRef = useRef(null)

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
      const res = await fetch('/Ejemplo Validacion de Nomina.xlsx')
      if (!res.ok) throw new Error('No se pudo cargar la plantilla de validación.')
      const templateBytes = new Uint8Array(await res.arrayBuffer())
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
          // cols 1(CODIGO EMPLEADO), 2(NOMBRES EMPLEADO), 3(CONCEPTO), 4(NOMBRE CONCEPTO)
          for (const col of [1, 2, 3, 4]) {
            const v = getCellValue(srcRow.getCell(col))
            if (v != null) { const s = String(v).trim(); if (s) neededStrings.add(s) }
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

        // ── Procesar rev incapacidades ───────────────────────────────────────
        const INCAP_CONCEPTOS = new Set(['001150', '001151', '001177', '001178'])
        const incapRows = nominaRows.filter(srcRow => {
          const concepto = String(getCellValue(srcRow.getCell(3)) ?? '').trim()
          return INCAP_CONCEPTOS.has(concepto)
        })
        const revPath = resolveSheetZipPath(workbookXml, relsXml, 'rev incapacidades')
        if (files[revPath] && incapRows.length > 0) {
          const revRowsXml = incapRows.map((srcRow, idx) => {
            const r = idx + 13 // datos desde fila 13 (encabezados en fila 12)
            return buildRevIncapRow(
              r,
              getCellValue(srcRow.getCell(1)), // CODIGO EMPLEADO
              getCellValue(srcRow.getCell(2)), // NOMBRES EMPLEADO
              getCellValue(srcRow.getCell(3)), // CONCEPTO
              getCellValue(srcRow.getCell(4)), // NOMBRE CONCEPTO
              getCellValue(srcRow.getCell(5)), // CANTIDAD
              getCellValue(srcRow.getCell(6)), // DEVENGOS
              getCellValue(srcRow.getCell(7)), // DEDUCCIONES
              strIdxMap
            )
          }).join('')
          files[revPath] = strToU8(rebuildSheetData(strFromU8(files[revPath]), revRowsXml, 12))

          // ── TablaDinámica4: calcular pivot P:S en "rev incapacidades" ──────
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
      }

      // ── Actualizar col A de "Seguridad Social" con códigos de empleados ──────
      if (fileMaestro && maestroRows.length > 0) {
        const ssPath = resolveSheetZipPath(workbookXml, relsXml, 'Seguridad Social')
        if (files[ssPath]) {
          const empCodes = maestroRows
            .map(row => String(getCellValue(row.getCell(2)) ?? '').trim())
            .filter(Boolean)
          files[ssPath] = strToU8(patchSegSocialColA(strFromU8(files[ssPath]), empCodes))
        }
      }

      // ── Generar archivos ─────────────────────────────────────────────────────
      delete files['xl/calcChain.xml']
      // Forzar recálculo al abrir en ambos archivos (Seguridad Social depende de Detallado Mes)
      files['xl/workbook.xml'] = strToU8(addFullCalcOnLoad(workbookXml))
      const fullBytes = zipSync(files)
      const segSocialBytes = await extractSegSocial(files, strFromU8(files['xl/workbook.xml']), relsXml)

      const parts = []
      if (file) parts.push(`${dataRows.length} registros IBC`)
      if (fileNomina) parts.push(`${nominaRows.length} registros nómina`)
      if (fileMaestro) parts.push(`${maestroRows.length} empleados maestro`)
      setModal({
        type: 'success',
        message: `Se procesaron ${parts.join(' y ')} correctamente.`,
        blobs: { full: fullBytes, segSocial: segSocialBytes },
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
                    <strong>Sube el archivo IBC</strong>
                    <p>El Excel de IBC del mes anterior.</p>
                  </div>
                </li>
                <li>
                  <span className="step-number">2</span>
                  <div>
                    <strong>Sube la nómina del mes</strong>
                    <p>El Excel de nómina detallada del mes actual.</p>
                  </div>
                </li>
                <li>
                  <span className="step-number">3</span>
                  <div>
                    <strong>Sube el Maestro Personal</strong>
                    <p>El Excel de Maestro Personal actualizado.</p>
                  </div>
                </li>
                <li>
                  <span className="step-number">4</span>
                  <div>
                    <strong>Genera y descarga</strong>
                    <p>Haz clic en <em>Generar Planilla de Validación</em>. Aparecerán dos opciones:</p>
                    <ul style={{marginTop: '6px', paddingLeft: '16px', lineHeight: '1.8'}}>
                      <li><strong>Libro completo</strong> — la plantilla entera con todas las hojas y los datos del mes cargados.</li>
                      <li><strong>Solo Seguridad Social</strong> — únicamente esa hoja con los valores ya calculados, lista para enviar o archivar.</li>
                    </ul>
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

                <button
                  className="btn-primary"
                  onClick={processFile}
                  disabled={(!file && !fileNomina && !fileMaestro) || isProcessing}
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

              </div>
            </div>
          </div>
        </div>
      </main>

      <footer className="footer">
        <div className="container">
          <p>&copy; {new Date().getFullYear()} Solutions & Payroll. Todos los derechos reservados.</p>
        </div>
      </footer>

      {modal && (
        <div className="modal-overlay" onClick={() => setModal(null)}>
          <div className={`modal-content ${modal.type}`} onClick={(e) => e.stopPropagation()}>
            <div className="modal-icon">
              {modal.type === 'success' ? (
                <svg width="50" height="50" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                  <polyline points="20 6 9 17 4 12"/>
                </svg>
              ) : (
                <svg width="50" height="50" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                  <circle cx="12" cy="12" r="10"/>
                  <line x1="12" y1="8" x2="12" y2="12"/>
                  <line x1="12" y1="16" x2="12.01" y2="16"/>
                </svg>
              )}
            </div>
            <h3 className="modal-title">{modal.type === 'success' ? '¡Listo!' : 'Error'}</h3>
            <p className="modal-message">{modal.message}</p>
            {modal.type === 'success' && modal.blobs ? (
              <div className="modal-downloads">
                <button
                  className="modal-download-btn primary"
                  onClick={() => downloadBlob(modal.blobs.segSocial, 'Seguridad_Social.xlsx')}
                >
                  <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                    <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
                    <polyline points="7 10 12 15 17 10"/>
                    <line x1="12" y1="15" x2="12" y2="3"/>
                  </svg>
                  Descargar Seguridad Social
                </button>
                <button
                  className="modal-download-btn secondary"
                  onClick={() => downloadBlob(modal.blobs.full, 'Validacion_Nomina.xlsx')}
                >
                  <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                    <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
                    <polyline points="7 10 12 15 17 10"/>
                    <line x1="12" y1="15" x2="12" y2="3"/>
                  </svg>
                  Descargar el libro de validación completo
                </button>
                <button className="modal-button" onClick={() => setModal(null)}>Cerrar</button>
              </div>
            ) : (
              <button className="modal-button" onClick={() => setModal(null)}>Cerrar</button>
            )}
            {modal.type === 'success' && !modal.blobs && (
              <div className="modal-progress">
                <div className="modal-progress-bar" onAnimationEnd={() => setModal(null)} />
              </div>
            )}
          </div>
        </div>
      )}
    </div>
  )
}

export default App
