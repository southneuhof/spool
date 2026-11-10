import ExcelJS from 'exceljs'
import { prepareDocument, type ExportStyles, type PageOrientation } from './print'
import { parseXlsxValue, type WorkbookCell, type WorkbookDocument, type WorkbookImage, type WorkbookSheet } from './workbook'

function positive(value: string | null, fallback: number): number {
  if (value === null) return fallback
  const number = Number(value)
  if (!Number.isFinite(number) || number < 0) throw new Error(`Invalid spreadsheet dimension: ${value}`)
  return number
}
function color(value: string): string | undefined {
  if (!value || value === 'transparent' || value === 'rgba(0, 0, 0, 0)') return undefined
  const hex = /^#([0-9a-f]{6})$/i.exec(value)
  if (hex?.[1]) return hex[1].toUpperCase()
  const rgb = /^rgba?\(\s*(\d+)[, ]+\s*(\d+)[, ]+\s*(\d+)(?:[, /]+([\d.]+))?\s*\)$/.exec(value)
  if (!rgb) return undefined
  return [rgb[1], rgb[2], rgb[3]].map(channel => Math.min(255, Number(channel)).toString(16).padStart(2, '0')).join('').toUpperCase()
}
function cellStyle(cell: HTMLTableCellElement): NonNullable<WorkbookCell['style']> {
  const css = cell.ownerDocument.defaultView?.getComputedStyle(cell)
  if (!css) return {}
  const alignment = css.textAlign === 'end' ? (css.direction === 'rtl' ? 'left' : 'right') : css.textAlign === 'start' ? (css.direction === 'rtl' ? 'right' : 'left') : css.textAlign
  const horizontal = alignment === 'center' || alignment === 'right' || alignment === 'justify' ? alignment : 'left'
  const vertical = css.verticalAlign === 'top' || css.verticalAlign === 'bottom' ? css.verticalAlign : 'middle'
  const borders: NonNullable<NonNullable<WorkbookCell['style']>['borders']> = {}
  for (const side of ['top', 'bottom', 'left', 'right'] satisfies ('top' | 'bottom' | 'left' | 'right')[]) {
    const style = css.getPropertyValue(`border-${side}-style`)
    const width = Number.parseFloat(css.getPropertyValue(`border-${side}-width`))
    const borderColor = color(css.getPropertyValue(`border-${side}-color`))
    if (style !== 'none' && width > 0 && borderColor) borders[side] = { color: borderColor, style: style === 'dashed' || style === 'dotted' || style === 'double' ? style : width >= 3 ? 'thick' : width >= 2 ? 'medium' : 'thin' }
  }
  const fontColor = color(css.color)
  let background = color(css.backgroundColor)
  let ancestor = cell.parentElement
  while (!background && ancestor && ancestor.tagName !== 'BODY' && ancestor.tagName !== 'HTML') {
    const parentCss = cell.ownerDocument.defaultView?.getComputedStyle(ancestor)
    if (parentCss) background = color(parentCss.backgroundColor)
    ancestor = ancestor.parentElement
  }
  return {
    fontName: css.fontFamily.split(',')[0]?.replace(/['"]/g, '') || 'Arial',
    fontSize: Number.parseFloat(css.fontSize) * 0.75 || 11,
    bold: css.fontWeight === 'bold' || Number(css.fontWeight) >= 600,
    italic: css.fontStyle === 'italic', horizontal, vertical,
    wrap: css.whiteSpace !== 'nowrap', borders,
    ...(fontColor ? { color: fontColor } : {}), ...(background ? { background } : {}),
  }
}
async function imageData(image: HTMLImageElement): Promise<Pick<WorkbookImage, 'data' | 'extension'>> {
  const response = await fetch(image.currentSrc || image.src)
  if (!response.ok) throw new Error(`Unable to load export image: ${image.src}`)
  const blob = await response.blob()
  const extension = blob.type === 'image/png' ? 'png' : blob.type === 'image/jpeg' ? 'jpeg' : blob.type === 'image/gif' ? 'gif' : undefined
  if (!extension) throw new Error(`Unsupported XLSX image format: ${blob.type}. Use PNG, JPEG, or GIF.`)
  const bytes = new Uint8Array(await blob.arrayBuffer())
  let binary = ''
  for (const byte of bytes) binary += String.fromCharCode(byte)
  return { data: `data:${blob.type};base64,${btoa(binary)}`, extension }
}

export async function extractWorkbook(target: HTMLElement): Promise<WorkbookDocument> {
  const document: WorkbookDocument = { sheets: [], diagnostics: [] }
  const wrappers = [...target.querySelectorAll<HTMLElement>('[data-xlsx-sheet]')].filter(element => !element.parentElement?.closest('[data-xlsx-sheet]') || element.parentElement?.closest('[data-xlsx-sheet]') === target)
  if (target.hasAttribute('data-xlsx-sheet')) wrappers.unshift(target)
  const roots = wrappers.length ? wrappers : [target]
  const names = new Set<string>()
  for (const [index, root] of roots.entries()) {
    const name = root.getAttribute('data-xlsx-sheet') || `Sheet${index + 1}`
    if (name.length > 31 || /[\\/*?:\[\]]/.test(name) || names.has(name.toLowerCase())) throw new Error(`Invalid or duplicate worksheet name: ${name}`)
    names.add(name.toLowerCase())
    const sheet: WorkbookSheet = { name, cells: [], columns: [], rows: [], images: [], tables: [], frozenRows: positive(root.getAttribute('data-xlsx-frozen-rows'), 0), frozenColumns: positive(root.getAttribute('data-xlsx-frozen-columns'), 0) }
    const tables = root.matches('table') ? [root] : [...root.querySelectorAll('table')].filter(table => (!table.parentElement?.closest('table') || !root.contains(table.parentElement.closest('table'))) && (wrappers.length === 0 || table.closest('[data-xlsx-sheet]') === root))
    let rowOffset = 0
    if (root.querySelector('table table')) throw new Error(`Nested HTML tables are not supported in worksheet ${name}. Use separate tables.`)
    for (const table of tables) {
      if (!(table instanceof (target.ownerDocument.defaultView?.HTMLTableElement ?? HTMLTableElement))) continue
      const occupied = new Set<string>()
      const startRow = rowOffset
      let columnCount = 0
      const columnElements = [...table.querySelectorAll('col')]
      const measuredCells: { element: HTMLTableCellElement; row: number; column: number; rowSpan: number; columnSpan: number }[] = []
      for (const [localRow, row] of [...table.rows].entries()) {
        const rowNumber = startRow + localRow + 1
        let column = 1
        let imageHeight = 0
        for (const element of [...row.cells]) {
          while (occupied.has(`${rowNumber}:${column}`)) column++
          const rowSpan = element.rowSpan || table.rows.length - localRow
          const columnSpan = element.colSpan
          measuredCells.push({ element, row: rowNumber, column, rowSpan, columnSpan })
          const raw = element.getAttribute('data-xlsx-value')
          const hyperlink = element.getAttribute('data-xlsx-link') ?? element.querySelector('a[href]')?.getAttribute('href')
          const format = element.getAttribute('data-xlsx-format')
          const css = element.ownerDocument.defaultView?.getComputedStyle(element)
          if (css && ((css.transform && css.transform !== 'none') || (css.backgroundImage && css.backgroundImage !== 'none') || (css.boxShadow && css.boxShadow !== 'none'))) document.diagnostics.push({ code: 'unsupported-css', message: 'Transforms, background images and shadows have no XLSX cell equivalent and were omitted.', sheet: name, cell: `${rowNumber}:${column}` })
          sheet.cells.push({ row: rowNumber, column, value: raw === null ? element.innerText ?? element.textContent ?? '' : parseXlsxValue(raw), rowSpan, columnSpan, style: cellStyle(element), ...(format ? { numberFormat: format } : {}), ...(hyperlink ? { hyperlink: new URL(hyperlink, element.baseURI).href } : {}) })
          for (let r = rowNumber; r < rowNumber + rowSpan; r++) for (let c = column; c < column + columnSpan; c++) {
            const key = `${r}:${c}`
            if (occupied.has(key)) throw new Error(`Overlapping merge in ${name} at row ${r}, column ${c}`)
            occupied.add(key)
          }
          const rect = element.getBoundingClientRect()
          const textWidth = Math.min(420, Math.max(64, (element.textContent?.length ?? 0) * 7 + 18))
          const measured = rect.width > 0 ? rect.width : textWidth
          for (let c = column; c < column + columnSpan; c++) {
            const col = columnElements[c - 1]
            const min = positive(col?.getAttribute('data-xlsx-min-width') ?? null, 8)
            const max = positive(col?.getAttribute('data-xlsx-max-width') ?? null, 60)
            const width = positive(col?.getAttribute('data-xlsx-width') ?? null, Math.min(max, Math.max(min, (measured / columnSpan - 5) / 7)))
            const old = sheet.columns[c - 1]
            const numFmt = col?.getAttribute('data-xlsx-format')
            sheet.columns[c - 1] = { width: Math.max(old?.width ?? 0, width), ...(numFmt ? { numberFormat: numFmt } : {}) }
          }
          for (const image of element.querySelectorAll('img')) {
            const geometry = image.getBoundingClientRect()
            imageHeight = Math.max(imageHeight, geometry.height || image.height || image.naturalHeight)
            sheet.images.push({ ...await imageData(image), row: rowNumber, column, width: geometry.width || image.width || image.naturalWidth, height: geometry.height || image.height || image.naturalHeight })
          }
          column += columnSpan
        }
        columnCount = Math.max(columnCount, column - 1)
        sheet.rows[rowNumber - 1] = { height: Math.max(15, row.getBoundingClientRect().height * 0.75, imageHeight * 0.75) }
      }
      const view = target.ownerDocument.defaultView
      if (view && 'CanvasRenderingContext2D' in view) {
        const context = target.ownerDocument.createElement('canvas').getContext('2d')
        for (const { element, row, column, rowSpan, columnSpan } of measuredCells) {
          if (!context) break
          const css = view.getComputedStyle(element)
          if (css.whiteSpace === 'nowrap') continue
          context.font = `${css.fontStyle} ${css.fontWeight} ${css.fontSize} ${css.fontFamily}`
          const padding = Number.parseFloat(css.paddingLeft || '0') + Number.parseFloat(css.paddingRight || '0')
          let width = -padding
          for (let c = column; c < column + columnSpan; c++) width += (sheet.columns[c - 1]?.width ?? 8) * 7 + 5
          let lines = 1
          let line = ''
          for (const word of (element.innerText ?? element.textContent ?? '').split(/\s+/)) {
            const next = line ? `${line} ${word}` : word
            if (line && context.measureText(next).width > Math.max(1, width)) { lines++; line = word } else line = next
          }
          const lineHeight = Number.parseFloat(css.lineHeight) || Number.parseFloat(css.fontSize) * 1.2
          const height = (lines * lineHeight + Number.parseFloat(css.paddingTop || '0') + Number.parseFloat(css.paddingBottom || '0')) * 0.75 / rowSpan
          for (let r = row; r < row + rowSpan; r++) sheet.rows[r - 1] = { height: Math.max(sheet.rows[r - 1]?.height ?? 15, height) }
        }
      }
      const tableName = table.getAttribute('data-xlsx-table')
      if (tableName) {
        if (sheet.cells.some(cell => cell.row > startRow && ((cell.rowSpan ?? 1) > 1 || (cell.columnSpan ?? 1) > 1))) throw new Error(`Native Excel table ${tableName} cannot contain merged cells.`)
        if (!/^[A-Za-z_][A-Za-z0-9_]*$/.test(tableName) || /^[A-Za-z]+\d+$/.test(tableName)) throw new Error(`Invalid native Excel table name: ${tableName}`)
        sheet.tables.push({ name: tableName, row: startRow + 1, column: 1, rowCount: table.rows.length, columnCount, filters: table.getAttribute('data-xlsx-filters') !== 'false' })
      }
      rowOffset += table.rows.length + 1
    }
    if (!sheet.cells.length) throw new Error(`Worksheet ${name} contains no table cells.`)
    sheet.rows = Array.from({ length: sheet.rows.length }, (_, index) => sheet.rows[index] ?? { height: 15 })
    document.sheets.push(sheet)
  }
  return document
}

export interface WriteXlsxOptions { orientation?: PageOrientation }
export async function writeXlsx(document: WorkbookDocument, options: WriteXlsxOptions = {}): Promise<Uint8Array> {
  const workbook = new ExcelJS.Workbook()
  workbook.calcProperties.fullCalcOnLoad = true
  for (const model of document.sheets) {
    const sheet = workbook.addWorksheet(model.name)
    if (options.orientation) sheet.pageSetup.orientation = options.orientation
    if (model.frozenRows || model.frozenColumns) sheet.views = [{ state: 'frozen', xSplit: model.frozenColumns, ySplit: model.frozenRows }]
    for (const [index, column] of model.columns.entries()) { sheet.getColumn(index + 1).width = column.width; if (column.numberFormat) sheet.getColumn(index + 1).numFmt = column.numberFormat }
    for (const [index, row] of model.rows.entries()) sheet.getRow(index + 1).height = row.height
    for (const cell of model.cells) {
      const output = sheet.getCell(cell.row, cell.column)
      output.value = cell.hyperlink ? { text: String(cell.value ?? ''), hyperlink: cell.hyperlink } : cell.value
      if (cell.numberFormat) output.numFmt = cell.numberFormat
      if (cell.style) {
        const style = cell.style
        output.font = { ...(style.fontName ? { name: style.fontName } : {}), ...(style.fontSize ? { size: style.fontSize } : {}), bold: style.bold ?? false, italic: style.italic ?? false, ...(style.color ? { color: { argb: `FF${style.color}` } } : {}) }
        output.alignment = { horizontal: style.horizontal ?? 'left', vertical: style.vertical ?? 'middle', wrapText: style.wrap ?? true }
        if (style.background) output.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: `FF${style.background}` } }
        for (const side of ['top', 'bottom', 'left', 'right'] satisfies ('top' | 'bottom' | 'left' | 'right')[]) {
          const border = style.borders?.[side]
          if (border) output.border = { ...output.border, [side]: { style: border.style, color: { argb: `FF${border.color}` } } }
        }
      }
      if ((cell.rowSpan ?? 1) > 1 || (cell.columnSpan ?? 1) > 1) sheet.mergeCells(cell.row, cell.column, cell.row + (cell.rowSpan ?? 1) - 1, cell.column + (cell.columnSpan ?? 1) - 1)
    }
    for (const table of model.tables) {
      const columns = Array.from({ length: table.columnCount }, (_, i) => ({ name: sheet.getCell(table.row, table.column + i).text, filterButton: table.filters }))
      if (new Set(columns.map(column => column.name)).size !== columns.length || columns.some(column => !column.name)) throw new Error(`Native Excel table ${table.name} requires unique, non-empty column headers.`)
      const rows = Array.from({ length: table.rowCount - 1 }, (_, r) => Array.from({ length: table.columnCount }, (_, c) => sheet.getCell(table.row + r + 1, table.column + c).value))
      sheet.addTable({ name: table.name, ref: sheet.getCell(table.row, table.column).address, headerRow: true, totalsRow: false, style: {
        // ExcelJS supports null to preserve authored cell styles; its published type omits it.
        // @ts-expect-error ExcelJS runtime explicitly supports a null table theme.
        theme: null, showRowStripes: false,
      }, columns, rows })
    }
    for (const image of model.images) {
      const id = workbook.addImage({ base64: image.data, extension: image.extension })
      sheet.addImage(id, { tl: { col: image.column - 1, row: image.row - 1 }, ext: { width: image.width, height: image.height }, editAs: 'oneCell' })
    }
  }
  return new Uint8Array(await workbook.xlsx.writeBuffer())
}
export interface DownloadXlsxOptions extends WriteXlsxOptions { filename?: string; styles?: ExportStyles }
export async function downloadXlsx(target: HTMLElement, options: DownloadXlsxOptions = {}): Promise<void> {
  const prepared = await prepareDocument({ target, layoutWidth: 1024, ...(options.styles ? { styles: options.styles } : {}) })
  try {
    const bytes = await writeXlsx(await extractWorkbook(prepared.target), options)
    const blob = new Blob([Uint8Array.from(bytes).buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' })
    const url = URL.createObjectURL(blob)
    const anchor = target.ownerDocument.createElement('a')
    anchor.href = url
    anchor.download = options.filename ?? 'export.xlsx'
    anchor.click()
    setTimeout(() => URL.revokeObjectURL(url), 1000)
  } finally { prepared.dispose() }
}
