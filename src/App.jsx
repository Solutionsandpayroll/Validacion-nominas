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

// ─── Extrae hoja "Seguridad Social" como workbook independiente ──────────────
// Fuerza recálculo completo al abrir en Excel
function addFullCalcOnLoad(wbXml) {
  if (wbXml.includes('<calcPr')) {
    // Agrega fullCalcOnLoad si no está ya
    return wbXml.replace(/<calcPr\b([^>]*?)\/>/,
      (_, attrs) => attrs.includes('fullCalcOnLoad')
        ? `<calcPr${attrs}/>`
        : `<calcPr${attrs} fullCalcOnLoad="1"/>`)
  }
  return wbXml.replace('</workbook>', '<calcPr fullCalcOnLoad="1"/></workbook>')
}

// Produce el ZIP "Seguridad Social standalone":
// Es el libro completo con todas las hojas de soporte (que tienen las fórmulas fuente),
// pero con todas ocultas excepto "Seguridad Social". Al abrirlo, Excel recalcula todo
// automáticamente y el usuario solo ve la hoja correcta con datos frescos.
function extractSegSocial(files, workbookXml, relsXml) {
  // Ocultar todas las hojas excepto "Seguridad Social"
  let newWbXml = workbookXml.replace(
    /<sheet\b([^>]*)\/>/g,
    (match, attrs) => {
      if (attrs.includes('name="Seguridad Social"')) {
        // Asegurarse de que no esté hidden
        return `<sheet${attrs.replace(/\s+state="[^"]*"/, '')}/>`
      }
      // Ocultar el resto (quitar state existente y poner hidden)
      return `<sheet${attrs.replace(/\s+state="[^"]*"/, '')} state="hidden"/>`
    }
  )
  // Forzar recálculo al abrir
  newWbXml = addFullCalcOnLoad(newWbXml)
  // Poner "Seguridad Social" como hoja activa (índice 18, base 0)
  newWbXml = newWbXml.includes('activeTab=')
    ? newWbXml.replace(/activeTab="\d+"/, 'activeTab="18"')
    : newWbXml.replace('<workbookView ', '<workbookView activeTab="18" ')

  const newFiles = { ...files, 'xl/workbook.xml': strToU8(newWbXml) }
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
  const [isHelpExpanded, setIsHelpExpanded] = useState(false)
  const [isProcessing, setIsProcessing] = useState(false)
  const [modal, setModal] = useState(null)
  const fileInputRef = useRef(null)
  const fileNominaRef = useRef(null)

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

  const removeFile = () => {
    setFile(null)
    if (fileInputRef.current) fileInputRef.current.value = ''
  }

  const removeFileNomina = () => {
    setFileNomina(null)
    if (fileNominaRef.current) fileNominaRef.current.value = ''
  }

  const formatBytes = (bytes) => {
    if (bytes < 1024) return bytes + ' B'
    if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB'
    return (bytes / (1024 * 1024)).toFixed(1) + ' MB'
  }

  const processFile = async () => {
    if (!file && !fileNomina) return
    setIsProcessing(true)
    try {
      let dataRows = []
      let nominaRows = []

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

      // ── Actualizar sharedStrings.xml ─────────────────────────────────────────
      const ssXml = strFromU8(files['xl/sharedStrings.xml'])
      const { strIdxMap, newSsXml } = buildSharedStringIndex(ssXml, neededStrings)
      files['xl/sharedStrings.xml'] = strToU8(newSsXml)

      // ── Procesar Planilla Mes anterior ───────────────────────────────────────
      if (file) {
        const sheetPath = resolveSheetZipPath(workbookXml, relsXml, 'Planilla Mes anterior')
        if (!files[sheetPath]) throw new Error(`No se encontró el archivo de hoja: ${sheetPath}`)

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
          return buildDetalladoRow(
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

        files[detPath] = strToU8(rebuildSheetData(strFromU8(files[detPath]), detRowsXml, 3))
      }

      // ── Generar archivos ─────────────────────────────────────────────────────
      delete files['xl/calcChain.xml']
      // Forzar recálculo al abrir en ambos archivos (Seguridad Social depende de Detallado Mes)
      files['xl/workbook.xml'] = strToU8(addFullCalcOnLoad(workbookXml))
      const fullBytes = zipSync(files)
      const segSocialBytes = extractSegSocial(files, strFromU8(files['xl/workbook.xml']), relsXml)

      const parts = []
      if (file) parts.push(`${dataRows.length} registros IBC`)
      if (fileNomina) parts.push(`${nominaRows.length} registros nómina`)
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
                    <p>Selecciona o arrastra el Excel de IBC del mes anterior (debe tener hoja "Sheet1")</p>
                  </div>
                </li>
                <li>
                  <span className="step-number">2</span>
                  <div>
                    <strong>Genera la planilla</strong>
                    <p>Haz clic en el botón para procesar y mapear los datos automáticamente</p>
                  </div>
                </li>
                <li>
                  <span className="step-number">3</span>
                  <div>
                    <strong>Descarga el resultado</strong>
                    <p>El Excel de validación se descargará con los datos en la hoja "Planilla Mes anterior"</p>
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

                <button
                  className="btn-primary"
                  onClick={processFile}
                  disabled={(!file && !fileNomina) || isProcessing}
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
