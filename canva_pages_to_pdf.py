#!/usr/bin/env python3
"""
Exporta un visor de Canva (u otro) a varios PDF, uno por página.
Avanza con la flecha derecha.

Uso:
  pip install selenium webdriver-manager
  python canva_pages_to_pdf.py "<url>" <carpeta_salida> <num_paginas> [ms_espera]

Ejemplo:
  python canva_pages_to_pdf.py "https://www.canva.com/design/.../view" ./canva_pdf 12 2000

Notas:
- Chrome visible para poder iniciar sesión; luego Enter en consola.
- Canva puede cambiar el comportamiento; ajusta tiempo o teclas si falla.
"""

from __future__ import annotations

import argparse
import base64
import sys
import time
from pathlib import Path

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager


def main() -> int:
    p = argparse.ArgumentParser(description="Canva / presentación → PDF por página (Selenium).")
    p.add_argument("url", help="URL del visor (p. ej. enlace view de Canva).")
    p.add_argument("out_dir", help="Carpeta de salida.")
    p.add_argument("pages", type=int, help="Número de páginas a capturar.")
    p.add_argument("delay_ms", type=int, nargs="?", default=2000, help="Espera antes de cada PDF (ms).")
    args = p.parse_args()

    if args.pages < 1:
        print("num_paginas debe ser >= 1", file=sys.stderr)
        return 1

    out = Path(args.out_dir).resolve()
    out.mkdir(parents=True, exist_ok=True)

    opts = Options()
    opts.add_argument("--window-size=1920,1080")
    opts.add_argument("--disable-blink-features=AutomationControlled")

    driver = webdriver.Chrome(
        service=Service(ChromeDriverManager().install()),
        options=opts,
    )
    try:
        driver.get(args.url)
        input(
            "\n>>> Inicia sesión si hace falta y deja la 1ª página visible.\n>>> Pulsa ENTER para empezar...\n"
        )

        body = driver.find_element(By.TAG_NAME, "body")

        for i in range(args.pages):
            time.sleep(max(0.5, args.delay_ms / 1000.0))
            pdf = driver.execute_cdp_cmd(
                "Page.printToPDF",
                {
                    "printBackground": True,
                    "paperWidth": 13.333,
                    "paperHeight": 7.5,
                    "marginTop": 0,
                    "marginBottom": 0,
                    "marginLeft": 0,
                    "marginRight": 0,
                },
            )
            raw = base64.b64decode(pdf["data"])
            path = out / f"pagina_{i + 1:03d}.pdf"
            path.write_bytes(raw)
            print(f"Guardado: {path}")
            if i < args.pages - 1:
                body.send_keys(Keys.RIGHT)

        print("Listo.")
        return 0
    finally:
        driver.quit()


if __name__ == "__main__":
    raise SystemExit(main())
