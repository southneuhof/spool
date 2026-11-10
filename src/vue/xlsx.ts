import { computed, defineComponent, h, ref, toValue, type MaybeRefOrGetter, type PropType } from 'vue'
import { downloadXlsx } from '../xlsx'
import { serializeXlsxValue, type XlsxValue } from '../workbook'
import type { ExportStyles, PageOrientation } from '../print'

export const XlsxSheet = defineComponent({
  name: 'XlsxSheet', inheritAttrs: false,
  props: { name: { type: String, required: true }, frozenRows: { type: Number, default: 0 }, frozenColumns: { type: Number, default: 0 } },
  setup: (props, { slots, attrs }) => () => h('section', { ...attrs, 'data-xlsx-sheet': props.name, 'data-xlsx-frozen-rows': props.frozenRows, 'data-xlsx-frozen-columns': props.frozenColumns, style: [attrs.style, { breakBefore: 'page' }] }, slots.default?.()),
})
export const XlsxCell = defineComponent({
  name: 'XlsxCell', inheritAttrs: false,
  props: { value: { type: [String, Number, Boolean, Date, Object] as PropType<XlsxValue>, default: undefined }, numberFormat: String, header: Boolean },
  setup: (props, { slots, attrs }) => () => h(props.header ? 'th' : 'td', { ...attrs, ...(props.value !== undefined ? { 'data-xlsx-value': serializeXlsxValue(props.value) } : {}), ...(props.numberFormat ? { 'data-xlsx-format': props.numberFormat } : {}) }, slots.default?.() ?? (props.value instanceof Date ? props.value.toLocaleDateString() : typeof props.value === 'object' && props.value !== null ? String(props.value.result ?? '') : String(props.value ?? ''))),
})
export const XlsxColumn = defineComponent({
  name: 'XlsxColumn', inheritAttrs: false,
  props: { width: Number, minWidth: Number, maxWidth: Number, numberFormat: String },
  setup: (props, { attrs }) => () => h('col', { ...attrs, 'data-xlsx-width': props.width, 'data-xlsx-min-width': props.minWidth, 'data-xlsx-max-width': props.maxWidth, 'data-xlsx-format': props.numberFormat }),
})
export const XlsxTable = defineComponent({
  name: 'XlsxTable', inheritAttrs: false,
  props: { name: String, filters: { type: Boolean, default: true } },
  setup: (props, { attrs, slots }) => () => h('table', { ...attrs, 'data-xlsx-table': props.name, 'data-xlsx-filters': String(props.filters) }, slots.default?.()),
})
export interface UseXlsxOptions {
  target: MaybeRefOrGetter<HTMLElement | null | undefined>
  filename?: MaybeRefOrGetter<string>
  styles?: ExportStyles
  orientation?: PageOrientation
}
export function useXlsx(options: UseXlsxOptions) {
  const pending = ref(false)
  const error = ref<Error | null>(null)
  let operation: Promise<void> | undefined
  const download = (): Promise<void> => {
    if (operation) return operation
    const target = toValue(options.target)
    if (!target) { error.value = new Error('The XLSX export target is not mounted.'); return Promise.reject(error.value) }
    pending.value = true
    error.value = null
    operation = downloadXlsx(target, { ...(options.orientation ? { orientation: options.orientation } : {}), ...(options.filename ? { filename: toValue(options.filename) } : {}), ...(options.styles ? { styles: options.styles } : {}) })
      .catch((cause: unknown) => { const failure = cause instanceof Error ? cause : new Error(String(cause)); error.value = failure; throw failure })
      .finally(() => { pending.value = false; operation = undefined })
    return operation
  }
  return { download, pending: computed(() => pending.value), error: computed(() => error.value) }
}
