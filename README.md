# Validación de Nóminas — William

Aplicación web para generar automáticamente la planilla de validación de nómina mensual a partir de tres archivos Excel fuente: **IBC**, **Nómina** y **Maestro Personal**.

---

## ¿Qué hace la app?

La aplicación toma la plantilla `Ejemplo Validacion de Nomina.xlsx` (25 hojas) y la rellena con los datos reales del mes sin modificar ninguna fórmula ni estructura. El resultado es un libro Excel listo para revisar en Excel con un solo clic en "Habilitar contenido".

### Hojas que se actualizan

| Hoja | Fuente | Qué se escribe |
|---|---|---|
| **Planilla Mes anterior** | Archivo IBC | Más de 80 columnas mapeadas desde Sheet1 + pivote de bases IBC en CX–CY |
| **Detallado Mes** | Archivo Nómina | Todas las novedades del mes con fórmulas de concatenación; pago SIP → HORAS = 0 |
| **rev incapacidades** | Archivo Nómina | Filas filtradas por conceptos 001150 / 001151 / 001177 / 001178 + tabla dinámica estática en P2:S15 |
| **Maestro Personal** | Archivo Maestro | Empleados activos/ausencia/vacaciones con tipos (fecha, número, texto) y estilos correctos |
| **Seguridad Social** | Maestro Personal | Columna A actualizada con los códigos de empleado vigentes; filas 362+ generadas si hay más de 359 empleados |

---

## Cómo usar la aplicación

### Paso 1 — Archivo IBC
Sube el Excel de IBC del mes **anterior**. Debe contener la hoja `Sheet1` con los datos de aportes del período. La app mapea más de 80 columnas y construye el pivote de bases IBC en CX–CY de la hoja *Planilla Mes anterior*.

### Paso 2 — Nómina del mes
Sube el Excel de nómina detallada del mes **actual**. Se extraen todas las novedades para *Detallado Mes*. Los conceptos de incapacidad se copian en *rev incapacidades* y se genera automáticamente su tabla de resumen por empleado.

### Paso 3 — Maestro Personal
Sube el Excel de Maestro Personal. Solo se toman los empleados con estado `Activos`, `Ausencia` o `Vacaciones` (fila 6 en adelante). Sus datos se vuelcan en la hoja *Maestro Personal* y sus códigos actualizan la columna A de *Seguridad Social*, de modo que todas las fórmulas de esa hoja apuntan al personal vigente del mes.

### Paso 4 — Generar
Haz clic en **Generar Planilla de Validación**. Se producen dos archivos descargables:

- **Libro completo** — las 25 hojas con datos cargados y fórmulas activas. Excel recalcula todo al abrirlo.
- **Solo Seguridad Social** — hoja independiente con valores ya evaluados (sin dependencias externas), lista para archivar o enviar.

> Los tres archivos son **opcionales** de forma individual. Puedes procesar solo IBC, solo nómina o cualquier combinación.

---

## Arquitectura técnica

```
src/
  App.jsx     — lógica completa (lectura, transformación, generación de XLSX)
  App.css     — estilos de la interfaz
public/
  Ejemplo Validacion de Nomina.xlsx   — plantilla base (no se modifica nunca)
```

### Librerías principales

| Librería | Versión | Rol |
|---|---|---|
| React 18 + Vite 5 | — | Interfaz y bundling |
| **ExcelJS** | 4.4.0 | Lee los tres archivos Excel fuente (nunca escribe) |
| **fflate** | — | Descomprime/recomprime el XLSX como ZIP para inyectar los datos |
| **HyperFormula** | gpl-v3 | Evalúa las fórmulas de *Seguridad Social* para el descargable standalone |
| vite-plugin-node-polyfills | — | Polyfills de Node.js requeridos por ExcelJS en el navegador |

### Cómo funciona internamente

1. **Lectura**: ExcelJS lee los tres archivos Excel fuente en el navegador (sin subir nada a ningún servidor).
2. **Plantilla**: fflate descomprime `Ejemplo Validacion de Nomina.xlsx` a nivel de ZIP, exponiendo los XML internos de cada hoja.
3. **Shared strings**: la app identifica todos los strings nuevos que necesita insertar y los añade al índice global (`xl/sharedStrings.xml`) antes de escribir las hojas.
4. **Escritura de hojas**: cada hoja se reescribe con `rebuildSheetData()`, que preserva las filas de encabezado originales y reemplaza los datos. Las celdas se construyen con el tipo, estilo y valor correctos (número, fecha, string compartido, fórmula).
5. **Seguridad Social**: `patchSegSocialColA()` actualiza solo la columna A de cada fila existente (3–361) y genera filas adicionales si hay más de 359 empleados, tomando la fila 3 como plantilla y renumerando todas las referencias de fila relativas.
6. **Standalone Seguridad Social**: `extractSegSocial()` usa HyperFormula para evaluar todas las fórmulas de la hoja con los datos reales y produce un XLSX independiente con valores estáticos.
7. **Recálculo**: se elimina `xl/calcChain.xml` y se fuerza `calcMode="auto" fullCalcOnLoad="1"` para que Excel recalcule todo al abrir.
8. **Descarga**: el ZIP final se genera en memoria con fflate y se descarga directamente desde el navegador.

### Rutas internas de las hojas clave

| Hoja | Ruta ZIP |
|---|---|
| Planilla Mes anterior | `xl/worksheets/sheet24.xml` |
| Detallado Mes | `xl/worksheets/sheet15.xml` |
| rev incapacidades | `xl/worksheets/sheet14.xml` |
| Maestro Personal | `xl/worksheets/sheet4.xml` |
| Seguridad Social | `xl/worksheets/sheet19.xml` |
| Parametros | `xl/worksheets/sheet10.xml` |

> Las rutas se resuelven dinámicamente desde `xl/workbook.xml` + `xl/_rels/workbook.xml.rels` usando `resolveSheetZipPath()`.

---

## Desarrollo local

```bash
npm install
npm run dev      # servidor de desarrollo en http://localhost:5173
npm run build    # build de producción en dist/
```

**Requisito**: Node.js 18+.

---

## Notas importantes

- **Todo el procesamiento es local** — ningún archivo se envía a un servidor. La app funciona completamente en el navegador.
- La plantilla `public/Ejemplo Validacion de Nomina.xlsx` **nunca se modifica**. Es la base de cada generación.
- Los archivos Excel fuente deben ser `.xlsx` (formato Office Open XML).
- El IBC debe tener sus datos en la hoja llamada exactamente `Sheet1`.
- El Maestro Personal se lee desde la fila 6 en adelante (las primeras 5 filas son encabezados/metadatos).
