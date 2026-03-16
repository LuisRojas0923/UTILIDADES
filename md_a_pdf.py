# -*- coding: utf-8 -*-
"""
Convierte Informe_Herramienta_Gestion_Nomina.md a HTML profesional
y luego a PDF usando Puppeteer (Chromium).
"""
import markdown
import subprocess
import os

BASE = r"c:\Users\amejoramiento6\Desktop\UTILIDADES"
MD_PATH = os.path.join(BASE, "Informe_Herramienta_Gestion_Nomina.md")
HTML_PATH = os.path.join(BASE, "Informe_Herramienta_Gestion_Nomina.html")
PDF_PATH = os.path.join(BASE, "Informe_Herramienta_Gestion_Nomina.pdf")
JS_PATH = os.path.join(BASE, "html_a_pdf.mjs")

with open(MD_PATH, "r", encoding="utf-8") as f:
    md_text = f.read()

html_body = markdown.markdown(md_text, extensions=["extra"])

html_doc = """<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="utf-8"/>
  <title>Informe — Herramienta de Gestión de Nómina</title>
  <style>
    @page { size: letter; margin: 0; }

    * { box-sizing: border-box; margin: 0; padding: 0; }

    body {
      font-family: "Segoe UI", Calibri, Arial, sans-serif;
      font-size: 10.5pt;
      line-height: 1.55;
      color: #222;
      padding: 2cm 2.5cm;
    }

    h1 {
      font-size: 17pt;
      font-weight: 700;
      color: #1b3a5c;
      border-bottom: 2.5px solid #1b3a5c;
      padding-bottom: 6px;
      margin-bottom: 14px;
    }

    .meta {
      background: #f0f4f8;
      border-left: 4px solid #1b3a5c;
      padding: 10px 14px;
      margin-bottom: 16px;
      font-size: 10pt;
      color: #333;
    }
    .meta p { margin: 3px 0; }

    h2 {
      font-size: 13pt;
      font-weight: 700;
      color: #1b3a5c;
      margin-top: 18px;
      margin-bottom: 8px;
      padding-bottom: 3px;
      border-bottom: 1px solid #d0d7de;
    }

    h3 {
      font-size: 11pt;
      font-weight: 600;
      color: #2d5986;
      margin-top: 14px;
      margin-bottom: 6px;
    }

    p {
      margin: 6px 0;
      text-align: justify;
    }

    ul, ol {
      margin: 6px 0 6px 22px;
      padding: 0;
    }

    li {
      margin: 4px 0;
      text-align: justify;
    }

    li > p {
      margin: 2px 0;
    }

    li > ul {
      margin-top: 2px;
      margin-bottom: 2px;
    }

    hr {
      border: none;
      border-top: 1px solid #c5cdd5;
      margin: 16px 0;
    }

    strong { font-weight: 600; }
    em { font-style: italic; color: #444; }

    .footer {
      margin-top: 20px;
      font-size: 9pt;
      font-style: italic;
      color: #666;
    }
  </style>
</head>
<body>
""" + html_body + """
</body>
</html>"""

# Envolver objetivo y alcance en bloque .meta
html_doc = html_doc.replace(
    '<p><strong>Objetivo del informe:</strong>',
    '<div class="meta">\n<p><strong>Objetivo del informe:</strong>'
)
html_doc = html_doc.replace(
    'sin atribuci\u00f3n de responsabilidades.</p>\n<hr',
    'sin atribuci\u00f3n de responsabilidades.</p>\n</div>\n<hr'
)

# Envolver pie de informe
html_doc = html_doc.replace(
    '<p><em>Documento generado con fines',
    '<p class="footer"><em>Documento generado con fines'
)

with open(HTML_PATH, "w", encoding="utf-8") as f:
    f.write(html_doc)
print("HTML generado:", HTML_PATH)

# Generar PDF con Puppeteer
result = subprocess.run(
    ["node", JS_PATH, HTML_PATH, PDF_PATH],
    capture_output=True, text=True, cwd=BASE
)
if result.returncode == 0:
    print("PDF generado:", PDF_PATH)
else:
    print("Error Puppeteer:", result.stderr)
