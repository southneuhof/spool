// @vitest-environment jsdom
import { describe, expect, it } from 'vitest'
import { preparePrint } from '../src/print'
import { exportCellText } from '../src/document'

describe('export boundaries', () => {
  it('rejects missing content without leaving an export frame', async () => {
    await expect(preparePrint({ target: document.createElement('article') })).rejects.toThrow('connected')
    expect(document.querySelectorAll('iframe')).toHaveLength(0)
  })
  it('does not start cancelled work', async () => {
    const target = document.createElement('article')
    document.body.append(target)
    const controller = new AbortController()
    controller.abort(new Error('Cancelled export'))
    await expect(preparePrint({ target, signal: controller.signal })).rejects.toThrow('Cancelled export')
    expect(document.querySelectorAll('iframe')).toHaveLength(0)
    target.remove()
  })
  it('keeps display text separate from typed spreadsheet values', () => {
    expect(exportCellText({ value: 1250, text: '$1,250' })).toBe('$1,250')
    expect(exportCellText({ value: { formula: 'SUM(A1:A2)', result: 3 } })).toBe('3')
    expect(exportCellText({})).toBe('')
  })
})
