/**
 * Exporta un visor de Canva (u otro) a varios PDF, uno por “página”.
 * Avanza con la flecha derecha (típico en presentación).
 *
 * Uso:
 *   node canva_pages_to_pdf.mjs "<url>" <carpeta_salida> <num_paginas> [ms_espera]
 *
 * Ejemplo:
 *   node canva_pages_to_pdf.mjs "https://www.canva.com/design/.../view" ./canva_pdf 12 2000
 *
 * Notas:
 * - Se abre Chrome visible: inicia sesión en Canva si lo pide, entra al diseño y pulsa ENTER en la consola.
 * - Canva puede cambiar el DOM; si no avanza, ajusta la tecla o el tiempo de espera.
 * - Respeta los términos de uso de Canva; uso bajo tu responsabilidad.
 */

import fs from "fs/promises";
import path from "path";
import puppeteer from "puppeteer";

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

const [, , urlArg, outDirArg, pagesArg, delayArg] = process.argv;
if (!urlArg || !outDirArg || !pagesArg) {
  console.error(
    "Uso: node canva_pages_to_pdf.mjs \"<url>\" <carpeta_salida> <num_paginas> [ms_espera]"
  );
  process.exit(1);
}

const url = urlArg;
const outDir = path.resolve(outDirArg);
const totalPages = Math.max(1, parseInt(pagesArg, 10) || 1);
const delayMs = Math.max(500, parseInt(delayArg || "2000", 10) || 2000);

await fs.mkdir(outDir, { recursive: true });

const browser = await puppeteer.launch({
  headless: false,
  channel: "chrome",
  executablePath: "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe",
  defaultViewport: { width: 1920, height: 1080 },
  args: [
    "--disable-blink-features=AutomationControlled",
    "--start-maximized",
  ],
});

const page = await browser.newPage();
await page.goto(url, { waitUntil: "domcontentloaded", timeout: 120_000 });

process.stdout.write(
  "\n>>> Inicia sesión si hace falta, abre la vista correcta y deja visible la 1ª página.\n>>> Pulsa ENTER aquí para empezar a guardar PDFs...\n"
);
await new Promise((resolve) => process.stdin.once("data", resolve));

for (let i = 0; i < totalPages; i++) {
  const n = String(i + 1).padStart(3, "0");
  const filePath = path.join(outDir, `pagina_${n}.pdf`);

  await sleep(delayMs);
  await page.waitForNetworkIdle({ idleTime: 500, timeout: 10_000 }).catch(() => {});

  await page.pdf({
    path: filePath,
    printBackground: true,
    width: "1920px",
    height: "1080px",
    margin: { top: "0", right: "0", bottom: "0", left: "0" },
  });
  console.log(`Guardado: ${filePath}  (${i + 1}/${totalPages})`);

  if (i < totalPages - 1) {
    await page.keyboard.press("ArrowRight");
    await sleep(1500);
  }
}

await browser.close();
console.log("Listo.");
