export type PageOrientation = 'portrait' | 'landscape'

export interface ExportStyles {
  mode: 'inherit' | 'isolated'
  stylesheets?: readonly string[]
  css?: string
}

export interface PreparePrintOptions {
  target: HTMLElement
  container?: HTMLElement
  styles?: ExportStyles
  title?: string
  signal?: AbortSignal
  timeoutMs?: number
  layoutWidth?: number
  orientation?: PageOrientation
}

export interface PreparedPrint {
  readonly document: Document
  readonly frame: HTMLIFrameElement
  readonly target: HTMLElement
  print(): void
  dispose(): void
}

const printCss = `
[data-spool-frame]{width:100%;border-collapse:collapse;table-layout:fixed;border:0}
[data-spool-frame]>thead{display:table-header-group}
[data-spool-frame]>tfoot{display:table-footer-group}
[data-spool-frame]>thead,[data-spool-frame]>tfoot{break-inside:avoid}
[data-spool-frame]>tbody>tr{break-inside:auto}
[data-spool-frame]>thead>tr>td,[data-spool-frame]>tbody>tr>td,[data-spool-frame]>tfoot>tr>td{padding:0;border:0;vertical-align:top}
[data-spool-footer]{position:fixed;bottom:0;left:0;right:0}
[data-print-break]{height:0;break-before:page;page-break-before:always}
[data-print-keep]{break-inside:avoid;page-break-inside:avoid}
html,body{margin:0;min-height:0}
[data-spool-context]{position:static!important;display:block!important;overflow:visible!important;transform:none!important;height:auto!important;min-height:0!important;max-height:none!important;width:auto!important}
*{-webkit-print-color-adjust:exact;print-color-adjust:exact}
@media print{
  h1,h2,h3,h4,h5,h6{break-after:avoid-page;page-break-after:avoid}
  p{orphans:3;widows:3}
  li,tr,figure,img{break-inside:avoid-page;page-break-inside:avoid}
  thead{display:table-header-group}
  tfoot{display:table-footer-group}
}
`

function copyAttributes(source: Element, destination: Element): void {
  for (const attribute of source.attributes) destination.setAttribute(attribute.name, attribute.value)
}

function group(document: Document, name: 'thead' | 'tbody' | 'tfoot', content: Node): HTMLElement {
  const section = document.createElement(name)
  const row = document.createElement('tr')
  const cell = document.createElement('td')
  cell.append(content)
  row.append(cell)
  section.append(row)
  return section
}

function frameHelpers(document: Document, target: HTMLElement): Node {
  const headers = target.querySelectorAll<HTMLElement>('[data-print-header="every-page"]')
  const footers = target.querySelectorAll<HTMLElement>('[data-print-footer="every-page"]')
  if (headers.length > 1 || footers.length > 1) throw new Error('Spool supports one repeated header and one repeated footer per report.')
  const header = headers.item(0)
  const footer = footers.item(0)
  if (!header && !footer) return target
  function contextFor(helper: HTMLElement): HTMLElement {
    let content = helper
    let ancestor = helper.parentElement
    while (ancestor) {
      const shell = document.importNode(ancestor, false)
      shell.removeAttribute('data-xlsx-sheet')
      shell.append(content)
      content = shell
      if (ancestor === target) break
      ancestor = ancestor.parentElement
    }
    return content
  }
  const headerContent = header ? contextFor(header) : null
  const footerContent = footer ? contextFor(footer) : null
  const wrapper = document.createElement('div')
  const table = document.createElement('table')
  table.setAttribute('data-spool-frame', '')
  table.setAttribute('role', 'presentation')
  if (headerContent) table.append(group(document, 'thead', headerContent))
  table.append(group(document, 'tbody', target))
  if (footerContent) {
    const reserve = document.importNode(footerContent, true)
    reserve.style.visibility = 'hidden'
    reserve.setAttribute('aria-hidden', 'true')
    table.append(group(document, 'tfoot', reserve))
    footerContent.setAttribute('data-spool-footer', '')
  }
  wrapper.append(table)
  if (footerContent) wrapper.append(footerContent)
  return wrapper
}

function copyLiveState(source: HTMLElement, clone: HTMLElement): HTMLElement {
  let result = clone
  const sources = [source, ...source.querySelectorAll('*')]
  const clones = [clone, ...clone.querySelectorAll('*')]
  for (const [index, original] of sources.entries()) {
    const destination = clones[index]
    if (!destination) continue
    if (original instanceof (source.ownerDocument.defaultView?.HTMLInputElement ?? HTMLInputElement) && 'checked' in destination && 'value' in destination) {
      destination.value = original.value
      destination.checked = original.checked
    } else if (original instanceof (source.ownerDocument.defaultView?.HTMLTextAreaElement ?? HTMLTextAreaElement) && 'value' in destination) {
      destination.value = original.value
      destination.textContent = original.value
    } else if (original instanceof (source.ownerDocument.defaultView?.HTMLSelectElement ?? HTMLSelectElement) && 'selectedIndex' in destination) {
      destination.selectedIndex = original.selectedIndex
    } else if (original instanceof (source.ownerDocument.defaultView?.HTMLCanvasElement ?? HTMLCanvasElement)) {
      const canvas = clone.ownerDocument.createElement('canvas')
      copyAttributes(original, canvas)
      canvas.width = original.width
      canvas.height = original.height
      canvas.getContext('2d')?.drawImage(original, 0, 0)
      if (destination === clone) result = canvas
      else destination.replaceWith(canvas)
    }
  }
  return result
}

function loaded(element: HTMLLinkElement | HTMLImageElement): Promise<void> {
  if ('naturalWidth' in element && element.complete) {
    return element.naturalWidth ? Promise.resolve() : Promise.reject(new Error(`Unable to load export image: ${element.src}`))
  }
  if ('href' in element && element.sheet) return Promise.resolve()
  return new Promise((resolve, reject) => {
    element.addEventListener('load', () => resolve(), { once: true })
    element.addEventListener('error', () => reject(new Error(`Unable to load export resource: ${'href' in element ? element.href : element.src}`)), { once: true })
  })
}

async function prepare(options: PreparePrintOptions, paginate: boolean): Promise<PreparedPrint> {
  const { target, signal } = options
  if (!target.isConnected) throw new Error('The export target must be connected to a document.')
  signal?.throwIfAborted()
  const source = target.ownerDocument
  const container = options.container ?? source.body
  if (!container.isConnected) throw new Error('The export container must be connected to a document.')
  const frame = container.ownerDocument.createElement('iframe')
  frame.title = options.title ?? 'Export'
  frame.setAttribute('aria-hidden', 'true')
  Object.assign(frame.style, { position: 'fixed', left: '-100000px', top: '0', width: `${options.layoutWidth ?? (target.getBoundingClientRect().width || 1024)}px`, height: '800px', border: '0' })
  container.append(frame)
  const document = frame.contentDocument
  const window = frame.contentWindow
  if (!document || !window) { frame.remove(); throw new Error('Unable to create the export document.') }
  let disposed = false
  let timer: ReturnType<typeof setTimeout> | undefined
  let rejectOnAbort: (() => void) | undefined
  const dispose = () => { if (!disposed) { disposed = true; frame.remove() } }
  const onAbort = () => dispose()
  signal?.addEventListener('abort', onAbort, { once: true })
  try {
    document.open()
    document.write('<!doctype html><html><head></head><body></body></html>')
    document.close()
    document.title = options.title ?? source.title
    const base = document.createElement('base')
    base.href = source.baseURI
    document.head.append(base)
    const styles = options.styles ?? { mode: 'inherit' }
    if (styles.mode === 'inherit') {
      copyAttributes(source.documentElement, document.documentElement)
      copyAttributes(source.body, document.body)
      for (const sheet of source.querySelectorAll('style,link[rel="stylesheet"]')) document.head.append(sheet.cloneNode(true))
    }
    for (const href of styles.stylesheets ?? []) {
      const link = document.createElement('link')
      link.rel = 'stylesheet'
      link.href = href
      document.head.append(link)
    }
    const css = document.createElement('style')
    css.textContent = `${printCss}\n${options.orientation ? `@page{size:${options.orientation}}` : ''}\n${styles.css ?? ''}`
    document.head.append(css)
    const clone = copyLiveState(target, document.importNode(target, true))
    const sheets = [...clone.querySelectorAll<HTMLElement>('[data-xlsx-sheet]')]
    if (clone.hasAttribute('data-xlsx-sheet')) sheets.unshift(clone)
    for (const [index, sheet] of sheets.entries()) sheet.style.breakBefore = index === 0 ? 'auto' : 'page'
    for (const script of clone.querySelectorAll('script')) script.remove()
    let content: Node = paginate ? frameHelpers(document, clone) : clone
    if (styles.mode === 'inherit') {
      let ancestor = target.parentElement
      while (ancestor && ancestor !== source.body) {
        const shell = document.createElement(ancestor.tagName)
        copyAttributes(ancestor, shell)
        shell.setAttribute('data-spool-context', '')
        shell.append(content)
        content = shell
        ancestor = ancestor.parentElement
      }
    }
    document.body.append(content)
    const ready = Promise.all([
      ...Array.from(document.querySelectorAll<HTMLLinkElement>('link[rel="stylesheet"]'), loaded),
      ...Array.from(document.images, image => { image.loading = 'eager'; return loaded(image) }),
    ]).then(async () => {
      await document.fonts.ready
      const failedFonts = [...document.fonts].filter(font => font.status === 'error')
      if (failedFonts.length) throw new Error(`Unable to load export fonts: ${failedFonts.map(font => font.family).join(', ')}`)
      await new Promise<void>(resolve => window.requestAnimationFrame(() => window.requestAnimationFrame(() => resolve())))
    })
    const interruption = new Promise<never>((_, reject) => {
      timer = setTimeout(() => reject(new Error('Export resources did not finish loading.')), options.timeoutMs ?? 30000)
      rejectOnAbort = () => reject(signal?.reason)
      signal?.addEventListener('abort', rejectOnAbort, { once: true })
    })
    await Promise.race([ready, interruption])
    signal?.throwIfAborted()
    return {
      document, frame, target: clone,
      print() {
        if (disposed) throw new Error('The prepared export has been disposed.')
        window.focus()
        window.print()
      },
      dispose,
    }
  } catch (error) { dispose(); throw error }
  finally {
    if (timer) clearTimeout(timer)
    signal?.removeEventListener('abort', onAbort)
    if (rejectOnAbort) signal?.removeEventListener('abort', rejectOnAbort)
  }
}

/** Prepare reusable content without moving print helpers into pagination wrappers. */
export function prepareDocument(options: PreparePrintOptions): Promise<PreparedPrint> {
  return prepare(options, false)
}

/** Prepare a document for native browser printing. */
export function preparePrint(options: PreparePrintOptions): Promise<PreparedPrint> {
  return prepare(options, true)
}

export async function print(options: PreparePrintOptions): Promise<void> {
  const prepared = await preparePrint(options)
  const window = prepared.frame.contentWindow
  window?.addEventListener('afterprint', prepared.dispose, { once: true })
  try { prepared.print() } catch (error) { prepared.dispose(); throw error }
}
