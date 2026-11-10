import type { XlsxValue } from './workbook'
export interface ExportCell { value?: XlsxValue; text?: string; numberFormat?: string; hyperlink?: string; class?: string; align?: 'start' | 'center' | 'end' }
export interface ExportColumn { id: string; label: string; class?: string }
export interface ExportRow { id?: string; cells: Record<string, ExportCell> }
export interface ExportTableDocument { kind: 'table'; title?: string; sheetName?: string; columns: readonly ExportColumn[]; rows: readonly ExportRow[] }
export interface ExportField extends ExportCell { id: string; label: string }
export interface ExportFieldsDocument { kind: 'fields'; title?: string; sheetName?: string; fields: readonly ExportField[] }
export type ExportDocument = ExportTableDocument | ExportFieldsDocument
export function exportCellText(cell: ExportCell): string {
  if (cell.text !== undefined) return cell.text
  const value = cell.value
  if (value === null || value === undefined) return ''
  if (value instanceof Date) return value.toLocaleDateString()
  if (typeof value === 'object') return value.result === undefined ? '' : String(value.result)
  return String(value)
}
