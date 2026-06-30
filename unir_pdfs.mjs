/**
 * Une todos los PDF de una carpeta en uno solo, ordenados alfabéticamente.
 *
 * Uso:
 *   node unir_pdfs.mjs <carpeta_pdfs> <archivo_salida.pdf>
 *
 * Ejemplo:
 *   node unir_pdfs.mjs .\canva_pdf .\Ricarpier_Detal_2026.pdf
 */

import fs from "fs/promises";
import path from "path";
import { PDFDocument } from "pdf-lib";

const [, , inputDir, outputFile] = process.argv;
if (!inputDir || !outputFile) {
  console.error("Uso: node unir_pdfs.mjs <carpeta_pdfs> <archivo_salida.pdf>");
  process.exit(1);
}

const dir = path.resolve(inputDir);
const files = (await fs.readdir(dir))
  .filter((f) => f.toLowerCase().endsWith(".pdf"))
  .sort();

if (files.length === 0) {
  console.error(`No se encontraron PDFs en ${dir}`);
  process.exit(1);
}

console.log(`Uniendo ${files.length} archivos PDF...`);

const merged = await PDFDocument.create();

for (const file of files) {
  const filePath = path.join(dir, file);
  const bytes = await fs.readFile(filePath);
  const doc = await PDFDocument.load(bytes);
  const pages = await merged.copyPages(doc, doc.getPageIndices());
  for (const p of pages) {
    merged.addPage(p);
  }
  console.log(`  + ${file}`);
}

const result = await merged.save();
await fs.writeFile(path.resolve(outputFile), result);
console.log(`\nPDF final: ${path.resolve(outputFile)} (${files.length} páginas)`);
