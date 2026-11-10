// @vitest-environment jsdom
import { createApp, h } from 'vue'
import { describe, expect, it } from 'vitest'
import { ExportTable } from '../src/vue/document'

describe('generic report authoring', () => {
  it('keeps arbitrary report titles separate from worksheet names and preserves custom rendered text', () => {
    const host = document.createElement('div')
    document.body.append(host)
    const app = createApp({ render: () => h(ExportTable, {
      document: {
        kind: 'table', title: 'A long report title / that is not a valid Excel worksheet name',
        columns: [{ id: 'status', label: 'Status', class: 'heading' }],
        rows: [{ cells: { status: { class: 'status', align: 'end' } } }],
      },
    }, { cell: () => h('span', 'Resolved status label') }) })
    app.mount(host)
    expect(host.querySelector('caption')?.textContent).toContain('A long report title /')
    expect(host.querySelector('table')?.getAttribute('data-xlsx-sheet')).toBe('Export')
    expect(host.querySelector('th')?.className).toBe('heading')
    expect(host.querySelector('td')?.className).toBe('status')
    expect(host.querySelector('td')?.style.textAlign).toBe('end')
    expect(host.querySelector('td')?.hasAttribute('data-xlsx-value')).toBe(false)
    expect(host.querySelector('td')?.textContent).toBe('Resolved status label')
    app.unmount()
    host.remove()
  })
})
