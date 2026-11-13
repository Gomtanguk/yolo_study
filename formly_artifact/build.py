import base64, pathlib, sys
root = pathlib.Path(__file__).resolve().parent
src = root / 'src'
css = (src / 'base.css').read_text() + '\n' + (src / 'extra.css').read_text()
# Sample workbook path is passed in (kept outside the repo: it contains real staff data).
sample_path = pathlib.Path(sys.argv[1]) if len(sys.argv) > 1 else src / 'travel_sample.xlsx'
sample = base64.b64encode(sample_path.read_bytes()).decode()
js = f"const SAMPLE_XLSX_B64 = '{sample}';\n" + (src / 'engine.js').read_text() + '\n' + (src / 'app.js').read_text()
assert '</script' not in js.lower()
html = f'''<!DOCTYPE html>
<html lang="ko"><head>
<meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>Formly · 문서 자동 작성</title>
<link rel="preconnect" href="https://fonts.googleapis.com"><link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Noto+Sans+KR:wght@400;500;700;800;900&display=swap" rel="stylesheet">
<style>{css}</style>
<script src="https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/hyperformula@2.7.1/dist/hyperformula.full.min.js"></script>
</head><body><div id="app"></div>
<script>
{js}
</script>
</body></html>
'''
(root / 'index.html').write_text(html)
print('index.html', len(html), 'bytes')
