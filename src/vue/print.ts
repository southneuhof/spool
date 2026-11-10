import { defineComponent, h, onScopeDispose, ref, shallowRef, toValue, type MaybeRefOrGetter, type PropType } from 'vue'
import { preparePrint, type ExportStyles, type PreparedPrint, type PageOrientation } from '../print'

export interface UsePrintOptions {
  target: MaybeRefOrGetter<HTMLElement | null | undefined>
  styles?: ExportStyles
  title?: string
  orientation?: PageOrientation
}

export function usePrint(options: UsePrintOptions) {
  const state = ref<'idle' | 'preparing' | 'ready' | 'printing' | 'error'>('idle')
  const error = shallowRef<unknown>(null)
  let current: PreparedPrint | undefined
  let controller: AbortController | undefined
  function dispose() { controller?.abort(); current?.dispose(); current = undefined; state.value = 'idle' }
  async function prepare() {
    dispose()
    controller = new AbortController()
    const active = controller
    state.value = 'preparing'
    error.value = null
    try {
      const target = toValue(options.target)
      if (!target) throw new Error('The export target is not mounted.')
      const result = await preparePrint({ target, signal: active.signal, ...(options.orientation ? { orientation: options.orientation } : {}), ...(options.styles ? { styles: options.styles } : {}), ...(options.title ? { title: options.title } : {}) })
      if (active.signal.aborted) { result.dispose(); active.signal.throwIfAborted() }
      current = result
      state.value = 'ready'
      return result
    } catch (cause) { if (!active.signal.aborted) { error.value = cause; state.value = 'error' }; throw cause }
  }
  async function print() {
    const prepared = await prepare()
    state.value = 'printing'
    prepared.frame.contentWindow?.addEventListener('afterprint', dispose, { once: true })
    try { prepared.print() } catch (cause) { dispose(); error.value = cause; state.value = 'error'; throw cause }
  }
  onScopeDispose(dispose)
  return { state, error, prepare, print, dispose }
}

export const PrintHeader = defineComponent({
  name: 'PrintHeader',
  props: { repeat: { type: String as PropType<'once' | 'every-page'>, default: 'once' } },
  setup(props, { slots }) { return () => h('header', { 'data-print-header': props.repeat }, slots.default?.()) },
})
export const PrintFooter = defineComponent({
  name: 'PrintFooter',
  props: { repeat: { type: String as PropType<'every-page'>, default: 'every-page' } },
  setup(props, { slots }) { return () => h('footer', { 'data-print-footer': props.repeat }, slots.default?.()) },
})
export const PrintPageBreak = defineComponent({
  name: 'PrintPageBreak',
  setup() { return () => h('div', { 'data-print-break': '', 'aria-hidden': 'true', style: { height: '0', breakBefore: 'page' } }) },
})
export const PrintKeepTogether = defineComponent({
  name: 'PrintKeepTogether',
  setup(_, { slots }) { return () => h('section', { 'data-print-keep': '', style: { breakInside: 'avoid' } }, slots.default?.()) },
})
