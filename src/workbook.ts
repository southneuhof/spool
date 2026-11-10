export type XlsxScalar = string | number | boolean | Date | null
export type XlsxValue = XlsxScalar | { formula: string; result?: string | number | boolean }
export interface WorkbookCell {
  row: number
  column: number
  value: XlsxValue
  rowSpan?: number
  columnSpan?: number
  numberFormat?: string
  hyperlink?: string
  style?: {
    fontName?: string; fontSize?: number; bold?: boolean; italic?: boolean
    color?: string; background?: string
    horizontal?: 'left' | 'center' | 'right' | 'justify'
    vertical?: 'top' | 'middle' | 'bottom'
    wrap?: boolean
    borders?: Partial<Record<'top' | 'bottom' | 'left' | 'right', { color: string; style: 'thin' | 'medium' | 'thick' | 'dashed' | 'dotted' | 'double' }>>
  }
}
export interface WorkbookImage {
  data: string
  extension: 'png' | 'jpeg' | 'gif'
  row: number
  column: number
  width: number
  height: number
}
export interface WorkbookSheet {
  name: string
  cells: WorkbookCell[]
  columns: { width: number; numberFormat?: string }[]
  rows: { height: number }[]
  images: WorkbookImage[]
  frozenRows: number
  frozenColumns: number
  tables: { name: string; row: number; column: number; rowCount: number; columnCount: number; filters: boolean }[]
}
export interface WorkbookDocument {
  sheets: WorkbookSheet[]
  diagnostics: { code: string; message: string; sheet: string; cell?: string }[]
}

export function serializeXlsxValue(value: XlsxValue): string {
  if (typeof value === 'number' && !Number.isFinite(value)) throw new Error('XLSX numeric values must be finite.')
  if (typeof value === 'object' && value !== null && !(value instanceof Date) && typeof value.result === 'number' && !Number.isFinite(value.result)) throw new Error('XLSX formula results must be finite.')
  return JSON.stringify(value instanceof Date ? { date: value.toISOString() } : value)
}

export function parseXlsxValue(input: string): XlsxValue {
  const value: unknown = JSON.parse(input)
  if (value === null || typeof value === 'string' || typeof value === 'boolean') return value
  if (typeof value === 'number' && Number.isFinite(value)) return value
  if (typeof value === 'object' && value !== null) {
    if ('date' in value && typeof value.date === 'string') {
      const date = new Date(value.date)
      if (Number.isFinite(date.getTime())) return date
    }
    if ('formula' in value && typeof value.formula === 'string' && value.formula.trim()) {
      if (!('result' in value) || value.result === undefined) return { formula: value.formula }
      if (typeof value.result === 'string' || typeof value.result === 'boolean' || (typeof value.result === 'number' && Number.isFinite(value.result))) return { formula: value.formula, result: value.result }
    }
  }
  throw new Error('Invalid data-xlsx-value. Expected a scalar, {date: ISO string}, or {formula, result?}.')
}
