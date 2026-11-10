import { defineComponent, h, type PropType } from 'vue'
import { exportCellText, type ExportCell, type ExportFieldsDocument, type ExportTableDocument } from '../document'
import { serializeXlsxValue } from '../workbook'
function metadata(cell: ExportCell) {
  return { class: cell.class, style: { textAlign: cell.align }, 'data-xlsx-value': cell.value === undefined ? undefined : serializeXlsxValue(cell.value), 'data-xlsx-format': cell.numberFormat, 'data-xlsx-link': cell.hyperlink }
}
export const ExportTable = defineComponent({
  name: 'ExportTable',
  props: { document: { type: Object as PropType<ExportTableDocument>, required: true } },
  setup(props, { slots }) {
    return () => h('table', { 'data-xlsx-sheet': props.document.sheetName ?? 'Export' }, [
      props.document.title ? h('caption', props.document.title) : null,
      h('thead', [h('tr', props.document.columns.map(column => h('th', { scope: 'col', key: column.id, class: column.class }, column.label)))]),
      h('tbody', props.document.rows.map((row, index) => h('tr', { key: row.id ?? index }, props.document.columns.map(column => {
        const cell = row.cells[column.id] ?? { value: null }
        return h('td', { key: column.id, ...metadata(cell) }, slots.cell?.({ row, column, cell, index }) ?? exportCellText(cell))
      })))),
    ])
  },
})
export const ExportFields = defineComponent({
  name: 'ExportFields',
  props: { document: { type: Object as PropType<ExportFieldsDocument>, required: true } },
  setup(props, { slots }) {
    return () => h('table', { 'data-xlsx-sheet': props.document.sheetName ?? 'Detail' }, [
      props.document.title ? h('caption', props.document.title) : null,
      h('tbody', props.document.fields.map((field, index) => h('tr', { key: field.id }, [
        h('th', { scope: 'row' }, field.label),
        h('td', metadata(field), slots.value?.({ field, index }) ?? exportCellText(field)),
      ]))),
    ])
  },
})
