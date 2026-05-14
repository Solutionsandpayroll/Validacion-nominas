# Validación de Nóminas — William

Aplicación web para generar automáticamente la planilla de validación de nómina mensual a partir de hasta seis archivos Excel fuente: **Plantilla base**, **IBC**, **Nómina**, **Nómina mes anterior**, **Maestro Personal** y **Novedades del cliente**.

---

## ¿Qué hace la app?

La aplicación toma la plantilla `Ejemplo Validacion de Nomina.xlsx` (25 hojas) y la rellena con los datos reales del mes sin modificar ninguna fórmula ni estructura. El resultado es un libro Excel listo para revisar con un solo clic en "Habilitar contenido".

### Hojas que se actualizan

| Hoja | Fuente | Qué se escribe |
|---|---|---|
| **Planilla Mes anterior** | Archivo IBC | Más de 80 columnas mapeadas desde Sheet1 + pivote de bases IBC en CX–CY |
| **Detallado Mes** | Archivo Nómina | Todas las novedades del mes con fórmulas de concatenación; pago SIP → HORAS = 0 |
| **rev incapacidades** | Archivo Nómina | Filas filtradas por conceptos 001150 / 001151 / 001177 / 001178; fórmula de días ajustada según porcentaje y referencia de parámetro por concepto |
| **Validacion novedades** | Nómina + Novedades | Cruce de novedades del período con BUSCARV a hoja Novedades; columnas I, J, K, L con lógica condicional por concepto |
| **Novedades** | Novedades del cliente | Conceptos y valores del período exportados desde *All Entitlements*; concepto 100015 multiplica su valor × 80.000 |
| **Maestro Personal** | Archivo Maestro | Empleados activos/ausencia/vacaciones con tipos (fecha, número, texto) y estilos correctos |
| **Seguridad Social** | Maestro Personal | Columna A actualizada con los códigos de empleado vigentes; filas 362+ generadas si hay más de 359 empleados |
| **validacion aux alimentacion (2)** | Maestro Personal + Nómina mes anterior | Empleados con sueldo, clase de salario, concepto de auxilio (111500/111501) según SMMLV y verificación contra el concepto del mes anterior |

---

## Cómo usar la aplicación

### Paso 1 — Plantilla base (opcional)
Si cuentas con una plantilla personalizada, súbela aquí. Si no, la app usa la plantilla interna `Ejemplo Validacion de Nomina.xlsx` por defecto.

### Paso 2 — IBC Mes Anterior
Sube el Excel de IBC del mes **anterior**. Debe contener la hoja `Sheet1` con los datos de aportes del período. La app mapea más de 80 columnas y construye el pivote de bases IBC en CX–CY de la hoja *Planilla Mes anterior*.

### Paso 3 — Archivo Nómina
Sube el Excel de nómina detallada del mes **actual**. Se extraen todas las novedades para *Detallado Mes*. Los conceptos de incapacidad (001150, 001151, 001177, 001178) se copian en *rev incapacidades* con su fórmula de días proporcional y porcentaje correspondiente. Adicionalmente se construye *Validacion novedades* con la liquidación cruzada contra las novedades del cliente.

### Paso 4 — Nómina mes anterior (opcional)
La nómina del mes pasado con la misma estructura. Se usa para determinar qué concepto de auxilio de alimentación (111500 ó 111501) tenía cada empleado el mes anterior y volcar ese valor estático en la columna H de *validacion aux alimentacion (2)*. Si no se sube, esa columna queda con una fórmula BUSCARV.

### Paso 5 — Maestro Personal
Sube el Excel de Maestro Personal actualizado. Solo se toman los empleados con estado `Activos`, `Ausencia` o `Vacaciones` (fila 6 en adelante). Sus datos llenan la hoja *Maestro Personal*, actualizan la columna A de *Seguridad Social* y construyen la hoja *validacion aux alimentacion (2)*.

### Paso 6 — Novedades del cliente (opcional)
El Excel de novedades exportado desde el sistema del cliente (*All Entitlements*). Se mapea a la hoja *Novedades* y alimenta el BUSCARV de la columna K en *Validacion novedades*.

### Paso 7 — Generar
Haz clic en **Generar**. Se descarga el **Libro completo** — las 25 hojas con datos cargados y fórmulas activas. Excel recalcula todo al abrirlo.

> Todos los archivos son **opcionales** de forma individual. Puedes procesar cualquier combinación según lo que tengas disponible.

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

1. **Lectura**: ExcelJS lee los archivos Excel fuente en el navegador (sin subir nada a ningún servidor).
2. **Plantilla**: fflate descomprime la plantilla `.xlsx` a nivel de ZIP, exponiendo los XML internos de cada hoja.
3. **Shared strings**: la app identifica todos los strings nuevos que necesita insertar y los añade al índice global (`xl/sharedStrings.xml`) antes de escribir las hojas.
4. **Escritura de hojas**: cada hoja se reescribe con `rebuildSheetData()`, que preserva las filas de encabezado originales y reemplaza los datos. Las celdas se construyen con el tipo, estilo y valor correctos (número, fecha, string compartido, fórmula).
5. **rev incapacidades**: se pre-construye un mapa `empCode → cantidad` del concepto 001150 para usarlo al generar las filas de 001177. Cada concepto recibe su propio porcentaje (70% u 100%) y su referencia de parámetro (`$C$3` a `$C$6`).
6. **Novedades**: se mapea desde la hoja *All Entitlements*; el concepto 100015 multiplica su valor numérico × 80.000.
7. **Validacion novedades**: columna J condicional según concepto (001150 → +F, 100015 → +I, resto → IF K≠0), columna L → IFERROR(J−K,0).
8. **validacion aux alimentacion (2)**: se construye desde Maestro Personal. Columna G determina el concepto de auxilio según SMMLV (`$H$1`). Columna H: si se subió nómina mes anterior, se escribe el concepto estático que tenía el empleado; si no, queda con BUSCARV.
9. **Seguridad Social**: `patchSegSocialColA()` actualiza solo la columna A de cada fila existente (3–361) y genera filas adicionales si hay más de 359 empleados, tomando la fila 3 como plantilla.
10. **Recálculo**: se elimina `xl/calcChain.xml` y se fuerza `calcMode="auto" fullCalcOnLoad="1"` para que Excel recalcule todo al abrir.
11. **Descarga**: el ZIP final se genera en memoria con fflate y se descarga directamente desde el navegador.

### Rutas internas de las hojas clave

| Hoja | Ruta ZIP |
|---|---|
| Planilla Mes anterior | `xl/worksheets/sheet24.xml` |
| Detallado Mes | `xl/worksheets/sheet15.xml` |
| rev incapacidades | `xl/worksheets/sheet14.xml` |
| Maestro Personal | `xl/worksheets/sheet4.xml` |
| Seguridad Social | `xl/worksheets/sheet19.xml` |
| validacion aux alimentacion (2) | `xl/worksheets/sheet3.xml` |
| Novedades | `xl/worksheets/sheet6.xml` |
| Validacion novedades | `xl/worksheets/sheet8.xml` |
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
- La plantilla `public/Ejemplo Validacion de Nomina.xlsx` **nunca se modifica**. Es la base de cada generación; si se sube una plantilla personalizada, se usa en su lugar.
- Los archivos Excel fuente deben ser `.xlsx` (formato Office Open XML).
- El IBC debe tener sus datos en la hoja llamada exactamente `Sheet1`.
- El Maestro Personal se lee desde la fila 6 en adelante (las primeras 5 filas son encabezados/metadatos).
- Las Novedades del cliente deben contener la hoja `All Entitlements` con los datos de novedades del período.
- La Nómina mes anterior debe tener la misma estructura de columnas que la nómina actual (código empleado en col A, concepto en col C).
