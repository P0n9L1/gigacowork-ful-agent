import pathlib, markdown
src = pathlib.Path("docs/REPORT.md").read_text(encoding="utf-8")
body = markdown.markdown(src, extensions=["tables", "fenced_code", "sane_lists"])
css = """
@page { size: A4; margin: 16mm 14mm; }
body { font-family: "DejaVu Sans", Arial, sans-serif; font-size: 9.4pt; line-height: 1.45; color: #1a1a1a; }
h1 { font-size: 19pt; color: #1F3864; border-bottom: 2px solid #1F3864; padding-bottom: 6px; margin-top: 0; }
h2 { font-size: 13pt; color: #1F3864; margin-top: 18px; page-break-after: avoid; }
h3 { font-size: 10.5pt; color: #2E5199; margin-top: 14px; page-break-after: avoid; }
table { border-collapse: collapse; width: 100%; margin: 8px 0 12px; font-size: 7.6pt; page-break-inside: auto; }
th { background: #1F3864; color: #fff; text-align: left; padding: 3px 4px; vertical-align: top; }
td { border: 0.4px solid #bbb; padding: 3px 4px; vertical-align: top; }
tr:nth-child(even) td { background: #F2F5FA; }
code { font-family: "DejaVu Sans Mono", monospace; font-size: 8.2pt; background: #f3f3f3; padding: 0 2px; }
pre { background: #f6f8fa; border-left: 3px solid #1F3864; padding: 7px 9px; font-size: 7.8pt; white-space: pre-wrap; page-break-inside: avoid; }
hr { border: none; border-top: 0.6px solid #ccc; margin: 14px 0; }
blockquote { color: #555; }
li { margin: 2px 0; }
"""
html = f"""<!doctype html><html lang="ru"><head><meta charset="utf-8">
<title>Агент распределения ФУЛ — отчёт</title><style>{css}</style></head><body>{body}</body></html>"""
pathlib.Path("docs/REPORT.html").write_text(html, encoding="utf-8")
print("html ok", len(html))
