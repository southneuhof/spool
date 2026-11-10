# Spreadsheet templates

Render a Vue component containing ordinary HTML tables. The same rendered component supplies the workbook preview, XLSX extraction, and browser PDF printing. Spreadsheet extraction reads table cells. Put titles and notes inside table rows when they must appear in Excel.

```vue
<script setup lang="ts">
import { XlsxSheet, XlsxCell, XlsxColumn } from 'spool/vue'
defineProps<{ rows: { name: string; amount: number }[] }>()
</script>
<template>
  <XlsxSheet name="Amounts" :frozen-rows="1">
    <table class="report-table">
      <colgroup>
        <XlsxColumn :min-width="12" :max-width="40" />
        <XlsxColumn number-format="#,##0.00" />
      </colgroup>
      <tr><th>Name</th><th>Amount</th></tr>
      <tr v-for="row in rows" :key="row.name">
        <td>{{ row.name }}</td>
        <XlsxCell :value="row.amount">{{ row.amount.toLocaleString() }}</XlsxCell>
      </tr>
    </table>
  </XlsxSheet>
</template>
```

`XlsxSheet` declares a worksheet. Names must be unique, contain at most 31 characters, and follow Excel's worksheet naming rules. Without sheet wrappers, all tables become one worksheet. Sibling tables have a blank separator row. Each sheet starts a new page when printed as PDF. All sheets are exported regardless of the preview's selected tab.

`XlsxCell` separates the stored value from its visible content. Values support strings, finite numbers, booleans, dates, null, and `{ formula, result? }`. Ordinary HTML cell contents remain strings, including leading zeroes and text beginning with `=`. Formulas require explicit metadata. `numberFormat` uses Excel format codes. `XlsxColumn` accepts width, minimum and maximum width in Excel character units. Automatic widths use browser measurements with bounds; wrapped row heights also account for the resulting column width. This approximates spreadsheet layout and does not promise pixel parity with HTML.

Optional `XlsxTable name="Records"` produces a native Excel table with filters. Its headers must be unique and nonempty, and its cells cannot be merged. Ordinary tables support `rowspan` and `colspan`. Frozen rows and columns belong to `XlsxSheet`.

Computed cell styles supply font family, size, weight, italic text, foreground and background colors, alignment, wrapping, and borders. Background colors can come from a parent row or table. CSS classes work through their computed styles. Arbitrary CSS has no complete spreadsheet equivalent. Unsupported transforms, background images, and shadows produce diagnostics in `WorkbookDocument.diagnostics`.

An ordinary anchor supplies a clickable hyperlink. Images inside cells support PNG, JPEG, and GIF. Remote images must permit browser fetching through CORS. Unsupported image formats reject export. Nested tables also reject export instead of silently losing content.

Plain HTML can use `data-xlsx-sheet`, `data-xlsx-value`, `data-xlsx-format`, and `data-xlsx-link` directly. The raw value attribute is JSON. Dates use `{ "date": "2026-11-04T00:00:00.000Z" }`. Use the Vue components or `serializeXlsxValue` to serialize values safely.

Core consumers can call `extractWorkbook(prepared.target)` to obtain the framework-neutral workbook model, inspect diagnostics, then call `writeXlsx(model)` for bytes. Prepare a connected document with `prepareDocument` first so resources and layout are ready, and dispose it afterward. `downloadXlsx(target, options)` handles preparation, serialization, download, and cleanup. Its measurement viewport is 1024 pixels wide.

For the real Chromium proof, serve this package through Vite on port 5184 and run `node tests/browser-xlsx/verify.cjs`. The fixture verifies two sheets, typed values, styles, merges, a link, an image, and stable sizing across two application viewport sizes. It then reopens the generated workbook with ExcelJS.

`writeXlsx(model, {orientation:'landscape'})`, `downloadXlsx`, and `useXlsx` can set the initial print orientation for all worksheets. This is workbook print metadata and does not change cell measurements.
