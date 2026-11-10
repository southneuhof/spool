# spool

Install with `npm install @southneuhof/spool`. Vue is an optional peer dependency for the Vue bindings.

Browser export library with a framework-independent browser core and Vue bindings. PDF uses Chromium's native print dialog and keeps text selectable. XLSX uses structured HTML tables and ExcelJS. No server is required.

```vue
<script setup lang="ts">
import { ref } from 'vue'
import { usePrint, useXlsx, PrintHeader, PrintFooter } from '@southneuhof/spool/vue'
const report = ref<HTMLElement | null>(null)
const { print } = usePrint({ target: report })
const { download } = useXlsx({ target: report, filename: 'report.xlsx' })
</script>
<template>
  <button @click="print">PDF</button>
  <button @click="download">Excel</button>
  <article ref="report">
    <PrintHeader repeat="every-page">Monthly report</PrintHeader>
    <table><tr><th>Name</th><th>Amount</th></tr><tr><td>Example</td><td>125</td></tr></table>
    <PrintFooter>Company report</PrintFooter>
  </article>
</template>
```

Render complete data before exporting. Vue components stay mounted in their normal application context. Preparation copies their resulting DOM into an export iframe and waits for stylesheets, images, and fonts. Export actions reject when resources fail or preparation times out. `usePrint` exposes `state`, `error`, `prepare`, `print`, and `dispose`. It cleans up on scope disposal and after the print dialog closes. Opening the print dialog does not confirm that a file was saved.

The default style mode inherits app stylesheets in order, ancestor selector context, HTML/body attributes, scoped CSS attributes, and theme variables. Ancestor layout containers are flattened for print so surrounding app UI is excluded. Isolated mode loads only supplied export styles and the structural helper rules. Inline styles inside the report remain.

```ts
usePrint({
  target: report,
  styles: { mode: 'isolated', stylesheets: ['/exports.css'], css: '.report { color: black }' },
})
```

The core exports `prepareDocument`, `preparePrint`, and `print`. `prepareDocument` creates a complete reusable snapshot for preview or later output. `preparePrint` additionally arranges repeated print helpers for pagination. Use the neutral document snapshot when preparing the content again. `preparePrint({target, styles?, title?, signal?, timeoutMs?, container?})` returns a snapshot with `target`, `document`, `frame`, `print()`, and idempotent `dispose()`. For a visible preview, pass a connected `container` when preparing and style the returned frame in place. Moving an initialized iframe into another container reloads its document in Chromium. Dispose snapshots you prepare manually. `print()` invokes the iframe's native print dialog. Paper, scale, and margins belong to that dialog. Optional `orientation: 'portrait' | 'landscape'` supplies a CSS page-orientation default without fixing the paper size. No PDF Blob or final page count is promised.

Optional Vue helpers are `PrintHeader` with `repeat="once"` (default) or `every-page`, `PrintFooter` with `repeat="every-page"`, `PrintPageBreak`, and `PrintKeepTogether`. One repeated header and footer are supported per report. Header and footer space is reserved on each page. A section taller than a page can split despite keep-together. The helpers do not set fonts, padding, paper, or margins.

`ExportTable` and `ExportFields` accept `ExportTableDocument` and `ExportFieldsDocument`. The `cell` slot receives `{row,column,cell,index}`. The `value` slot receives `{field,index}`. Columns accept header classes. Cells accept classes and logical text alignment. Report titles remain display text; optional `sheetName` controls the worksheet name, which defaults to `Export` or `Detail`. Typed cell values and displayed text are separate. Omitting `value` lets XLSX read a custom renderer's visible text. Include spreadsheet titles and notes as table rows if they belong in the workbook.

XLSX authoring details and supported limitations are in [XLSX.md](XLSX.md).

## Verification

Run `tsc --noEmit` and `vitest run` from this package. Browser proof fixtures live in `tests/browser`. Start Vite on port 5184 from this package, then run `node tests/browser/verify.cjs`. Set `CHROMIUM_PATH` when Chromium is installed elsewhere. The script verifies native print invocation, inherited and isolated colors, failed resource cleanup, disposal, and writes `tests/browser/report.pdf` from the prepared document for native PDF text and pagination inspection. Run `python3 tests/browser/verify_pdf.py` with PyMuPDF installed to assert all 90 native text rows and repeated helper geometry. Browser-dialog interaction itself is manual because headless Chromium exposes PDF bytes through its print engine.


## Export your own component without Loom

Render your custom component inside a connected target element. Call the Vue
bindings directly. This bypasses resource configuration and Loom's default
preview and templates.

```vue
<script setup lang="ts">
import { ref } from 'vue'
import { usePrint, useXlsx } from '@southneuhof/spool/vue'
import MyExport from './MyExport.vue'
const props = defineProps<{ rows: { name: string; amount: number }[] }>()
const target = ref<HTMLElement | null>(null)
const { print } = usePrint({ target })
const { download } = useXlsx({ target, filename: 'my-export.xlsx' })
</script>
<template>
	<button @click="print">Print / Save PDF</button>
	<button @click="download">Download Excel</button>
	<div ref="target"><MyExport :rows="props.rows" /></div>
</template>
```

The rendered `MyExport` is the preview. For PDF it can contain any report HTML
and print helpers. For Excel it needs HTML tables, optionally composed with
`XlsxSheet`, `XlsxCell`, and `XlsxColumn`. Both bindings support isolated styling
through their `styles` option; otherwise your component's app styles carry over.
You own data authorization and loading outside Loom. Declare `spool` as a direct
workspace dependency in any package that imports these bindings.

Both `usePrint` and `useXlsx` accept an `orientation` option. PDF uses a native `@page` rule. XLSX writes the orientation into each worksheet's print setup. Explicit custom `@page` CSS can override the PDF default.

## Native print headers and pagination

Chromium's timestamp, document title, URL, and page numbers belong to its print
settings. `window.print()` cannot switch them off. Clear **More settings → Headers
and footers** in the print dialog. Spool does not force zero page margins to
suppress these decorations, because margins remain a native print setting.

Print preparation asks Chromium to keep headings with the following content,
keep table rows, list items, images, and figures together, repeat table headers,
and avoid single-line paragraph fragments. For a custom group, wrap its related
content in `PrintKeepTogether`. Group each history entry, not the whole history
section. A group taller than the printable page must still split. Custom print
CSS can override these defaults.
