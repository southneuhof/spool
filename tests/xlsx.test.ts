import { describe, expect, it } from 'vitest'
import ExcelJS from 'exceljs'
import { writeXlsx } from '../src/xlsx'
import { parseXlsxValue, serializeXlsxValue, type WorkbookDocument } from '../src/workbook'

describe('workbook output', () => {
  it('roundtrips typed values, layout, merges, links, images and multiple worksheets through the real XLSX archive', async () => {
    const date = new Date('2026-11-04T00:00:00Z')
    const model: WorkbookDocument = { diagnostics: [], sheets: [{ name: 'People', frozenRows: 1, frozenColumns: 1, columns: [{ width: 24 }, { width: 30 }], rows: [{ height: 30 }], tables: [], images: [{ data: 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAIAAACQkWg2AAAAGUlEQVR4nGNkqDjBQApgIkn1qIZRDUNKAwAFBAFgtwPMAwAAAABJRU5ErkJggg==', extension: 'png', row: 1, column: 2, width: 20, height: 20 }], cells: [
      { row: 1, column: 1, value: 'Title', columnSpan: 2, style: { bold: true, background: 'FF0000', color: 'FFFFFF', wrap: true } },
      { row: 2, column: 1, value: '00123' }, { row: 2, column: 2, value: 23.5, numberFormat: '#,##0.00' },
      { row: 3, column: 1, value: date, numberFormat: 'yyyy-mm-dd' }, { row: 3, column: 2, value: true },
      { row: 4, column: 1, value: { formula: 'SUM(B2,1)', result: 24.5 } },
      { row: 4, column: 2, value: 'Visit', hyperlink: 'https://example.com' },
    ] }, { name: 'Other', frozenRows: 0, frozenColumns: 0, columns: [], rows: [], images: [], tables: [], cells: [{ row: 1, column: 1, value: '=not a formula' }] }] }
    const bytes = await writeXlsx(model, { orientation: 'landscape' })
    expect(bytes[0]).toBe(80)
    expect(bytes[1]).toBe(75)
    const loaded = new ExcelJS.Workbook()
    await loaded.xlsx.load(Uint8Array.from(bytes).buffer)
    expect(loaded.worksheets.map(sheet => sheet.name)).toEqual(['People', 'Other'])
    expect(loaded.worksheets.every(sheet => sheet.pageSetup.orientation === 'landscape')).toBe(true)
    const sheet = loaded.getWorksheet('People')!
    expect(sheet.getCell('A2').value).toBe('00123')
    expect(sheet.getCell('B2').value).toBe(23.5)
    expect(sheet.getCell('B2').numFmt).toBe('#,##0.00')
    expect(sheet.getCell('A3').value).toEqual(date)
    expect(sheet.getCell('B3').value).toBe(true)
    expect(sheet.getCell('A4').value).toEqual({ formula: 'SUM(B2,1)', result: 24.5 })
    expect(sheet.getCell('B4').value).toEqual({ text: 'Visit', hyperlink: 'https://example.com' })
    expect(sheet.getCell('B1').master.address).toBe('A1')
    expect(sheet.getCell('A1').font.bold).toBe(true)
    expect(sheet.getColumn(1).width).toBe(24)
    expect(sheet.getRow(1).height).toBe(30)
    expect(sheet.views[0]).toMatchObject({ state: 'frozen', xSplit: 1, ySplit: 1 })
    expect(sheet.getImages()).toHaveLength(1)
    expect(loaded.getWorksheet('Other')?.getCell('A1').value).toBe('=not a formula')
  })
  it('retains explicit raw metadata without guessing displayed text types', () => {
    expect(parseXlsxValue(serializeXlsxValue('00123'))).toBe('00123')
    expect(parseXlsxValue(serializeXlsxValue(new Date('2026-11-04')))).toEqual(new Date('2026-11-04'))
    expect(() => parseXlsxValue('{"formula":3}')).toThrow('Invalid data-xlsx-value')
  })
})
