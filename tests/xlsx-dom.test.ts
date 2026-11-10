// @vitest-environment jsdom
import { describe, expect, it } from 'vitest'
import ExcelJS from 'exceljs'
import { extractWorkbook, writeXlsx } from '../src/xlsx'

function fixture(html: string): HTMLElement {
  const root = document.createElement('div')
  root.innerHTML = html
  document.body.append(root)
  return root
}
describe('HTML spreadsheet authoring', () => {
  it('exports all sheets, ordinary HTML strings, explicit raw values, CSS, merges and links', async () => {
    const root = fixture(`<section data-xlsx-sheet="First" data-xlsx-frozen-rows="1"><table><colgroup><col data-xlsx-width="25"><col data-xlsx-format="0.00"></colgroup><tr><th colspan="2">Heading</th></tr><tr><td style="color:rgb(255,0,0);background-color:rgb(0,255,0);font-weight:700">00123</td><td data-xlsx-value="42" data-xlsx-format="0.0">Forty two</td></tr><tr><td><a href="https://example.com">Website</a></td><td data-xlsx-value='{"formula":"B2*2","result":84}'>84</td></tr></table></section><section data-xlsx-sheet="Second" hidden><table><tr><td>true</td></tr></table></section>`)
    const model = await extractWorkbook(root)
    expect(model.sheets).toHaveLength(2)
    expect(model.sheets[0]?.columns[0]?.width).toBe(25)
    expect(model.sheets[0]?.cells[1]).toMatchObject({ value: '00123', style: { color: 'FF0000', background: '00FF00', bold: true } })
    const loaded = new ExcelJS.Workbook()
    await loaded.xlsx.load(Uint8Array.from(await writeXlsx(model)).buffer)
    expect(loaded.getWorksheet('First')?.getCell('B2').value).toBe(42)
    expect(loaded.getWorksheet('First')?.getCell('A3').value).toEqual({ text: 'Website', hyperlink: 'https://example.com/' })
    expect(loaded.getWorksheet('Second')?.getCell('A1').value).toBe('true')
    root.remove()
  })
  it('writes requested named native Excel tables with filters and rejects merged table cells', async () => {
    const root = fixture('<table data-xlsx-table="Records"><tr><th>Name</th><th>Count</th></tr><tr><td>A</td><td data-xlsx-value="3">3</td></tr></table>')
    const loaded = new ExcelJS.Workbook()
    await loaded.xlsx.load(Uint8Array.from(await writeXlsx(await extractWorkbook(root))).buffer)
    expect(loaded.getWorksheet('Sheet1')?.getTable('Records')).toBeDefined()
    expect(loaded.getWorksheet('Sheet1')?.getCell('B2').value).toBe(3)
    root.innerHTML = '<table data-xlsx-table="Records"><tr><th colspan="2">Merged</th></tr></table>'
    await expect(extractWorkbook(root)).rejects.toThrow('cannot contain merged cells')
    root.remove()
  })
  it('does not turn the page background into cell fills', async () => {
    document.body.style.backgroundColor = 'white'
    const root = fixture('<table><tr><th style="border:1px solid black">Name</th></tr><tr><td>Plain</td></tr></table>')
    try {
      const model = await extractWorkbook(root)
      expect(model.sheets[0]?.cells.every(cell => cell.style?.background === undefined)).toBe(true)
      const loaded = new ExcelJS.Workbook()
      await loaded.xlsx.load(Uint8Array.from(await writeXlsx(model)).buffer)
      expect(loaded.getWorksheet('Sheet1')?.getCell('A2').fill).toEqual({ type: 'pattern', pattern: 'none' })
    } finally { root.remove(); document.body.style.backgroundColor = '' }
  })
  it('maps logical alignment to the spreadsheet cell direction', async () => {
    const root = fixture('<table><tr><td style="text-align:end;direction:ltr">Amount</td><td style="text-align:start;direction:rtl">RTL label</td></tr></table>')
    const model = await extractWorkbook(root)
    expect(model.sheets[0]?.cells.map(cell => cell.style?.horizontal)).toEqual(['right', 'right'])
    root.remove()
  })
  it('keeps a blank separator between sibling tables when writing a worksheet', async () => {
    const root = fixture('<table><tr><td>First table</td></tr></table><table><tr><td>Second table</td></tr></table>')
    const loaded = new ExcelJS.Workbook()
    await loaded.xlsx.load(Uint8Array.from(await writeXlsx(await extractWorkbook(root))).buffer)
    const sheet = loaded.getWorksheet('Sheet1')
    expect(sheet?.getCell('A1').value).toBe('First table')
    expect(sheet?.getCell('A2').value).toBeNull()
    expect(sheet?.getCell('A3').value).toBe('Second table')
    root.remove()
  })
  it('rejects invalid names and nested tables rather than dropping content', async () => {
    const root = fixture('<section data-xlsx-sheet="Invalid/Name"><table><tr><td>A</td></tr></table></section>')
    await expect(extractWorkbook(root)).rejects.toThrow('Invalid or duplicate worksheet name')
    root.innerHTML = '<table><tr><td>A<table><tr><td>Nested</td></tr></table></td></tr></table>'
    await expect(extractWorkbook(root)).rejects.toThrow('Nested HTML tables')
    root.remove()
  })
})
