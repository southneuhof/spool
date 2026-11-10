"""Check native text and helper geometry in Chromium's actual PDF output."""
import json
from pathlib import Path
import fitz

report = fitz.open(Path(__file__).with_name('report.pdf'))
row_count = 0
for number, page in enumerate(report):
    headers = page.search_for('Spool repeated header')
    footers = page.search_for('Spool repeated footer')
    assert len(headers) == 1, (number, 'header')
    assert len(footers) == 1, (number, 'footer')
    assert footers[0].y0 > page.rect.height - 100, (number, 'bottom anchored')
    rows = page.search_for('Selectable export content')
    assert all(row.y0 > headers[0].y1 and row.y1 < footers[0].y0 for row in rows), (number, 'overlap')
    row_count += len(rows)
assert row_count == 90, row_count
assert any('Together heading' in page.get_text() and 'Together content' in page.get_text() for page in report)
print(json.dumps({'pages': len(report), 'nativeRows': row_count, 'repeatedHelpers': True, 'noOverlap': True}))
