import { createApp, defineComponent, h } from 'vue'
import { XlsxCell, XlsxSheet } from '../../src/vue/xlsx'
import { prepareDocument } from '../../src/print'
import { extractWorkbook, writeXlsx } from '../../src/xlsx'
const image = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAIAAACQkWg2AAAAGUlEQVR4nGNkqDjBQApgIkn1qIZRDUNKAwAFBAFgtwPMAwAAAABJRU5ErkJggg=='
const Fixture = defineComponent({ setup() { return () => h('article', { id: 'report' }, [
  h(XlsxSheet, { name: 'First', frozenRows: 1 }, () => h('table', { class: 'export-table' }, [
    h('tr', [h('th', { colspan: 2 }, 'Spool browser workbook')]),
    h('tr', [h('td', { class: 'important' }, '00123'), h(XlsxCell, { value: 42.5, numberFormat: '0.00' }, () => 'Forty two and a half')]),
    h('tr', { style: { backgroundColor: 'rgb(220, 230, 240)' } }, [h(XlsxCell, { value: new Date('2026-11-04T00:00:00Z'), numberFormat: 'yyyy-mm-dd' }, () => '4 November 2026'), h(XlsxCell, { value: { formula: 'B2*2', result: 85 } }, () => '85')]),
    h('tr', [h('td', [h('a', { href: 'https://example.com' }, 'Website')]), h('td', [h('img', { src: image, width: 80, height: 40 })])]),
    h('tr', [h('td', 'A long description wraps within the bounded column instead of creating an excessively wide worksheet column.'), h(XlsxCell, { value: true }, () => 'Yes')]),
  ])),
  h(XlsxSheet, { name: 'Second' }, () => h('table', { class: 'export-table detail' }, [h('tr', [h('td', 'Still included')])])),
]) } })
createApp(Fixture).mount('#app')
Object.assign(window, { xlsxProof: async () => {
  const target = document.getElementById('report')
  if (!target) throw new Error('Missing fixture')
  const prepared = await prepareDocument({ target, layoutWidth: 1024 })
  try {
    const model = await extractWorkbook(prepared.target)
    const geometry = [...prepared.target.querySelectorAll('table')].map(table => table.getBoundingClientRect().width)
    const bytes = await writeXlsx(model)
    return { model, geometry, bytes: Array.from(bytes) }
  } finally { prepared.dispose() }
} })
