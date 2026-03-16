import java.io.FileOutputStream;
import java.io.File;
import java.awt.Color;
import java.awt.Toolkit;
import java.sql.ResultSet;
import java.util.TreeMap;
import java.util.TreeSet;
import java.util.Map;
import java.util.ArrayList;
import java.util.List;
import javax.swing.JOptionPane;
import com.lowagie.text.*;
import com.lowagie.text.pdf.*;

/**
 * Generador de Reporte de Gastos de Viáticos siguiendo el diseño Refridcol.
 */
public class ReporteGastosViaticos extends BaseInforme {

    private ConexionDatos conexion;
    private String dirArchivo;
    private java.awt.Image imageLogo = null;

    public ReporteGastosViaticos(ConexionDatos inConexion) {
        super();
        this.conexion = inConexion;
        cargarLogo();
    }

    private void cargarLogo() {
        try {
            // Se prioriza el logo local suministrado por el usuario
            String pathLocal = "logoBarra1.png";
            File archivoLogo = new File(pathLocal);
            if (archivoLogo.exists()) {
                imageLogo = Toolkit.getDefaultToolkit().getImage(archivoLogo.getAbsolutePath());
            } else {
                // Si no existe, intenta cargar desde BD como respaldo
                Imagen tmpImagen = new Imagen(null, conexion);
                tmpImagen.cargar("codigo", "3");
                imageLogo = ImageConverter.convertToImage(tmpImagen.representacion);
            }
        } catch (Exception e) {
            // Logo vacío si falla todo
        }
    }

    @Override
    public void generar(String codigoLegalizacion) throws Exception {
        // Consultar cabecera
        String sqlL = "SELECT * FROM legalizacion WHERE codigolegalizacion = '" + codigoLegalizacion + "'";
        ResultSet rsL = conexion.funcionConsultar(sqlL);
        String[][] datosL = ConexionDatos.armarArreglo(rsL);

        if (datosL == null || datosL[0].length < 2) {
            throw new Exception("No se encontró la legalización: " + codigoLegalizacion);
        }

        // --- DINAMIZAR EL CAMPO OBRA (CLIENTE) ---
        String clienteEncontrado = "-";
        try {
            // 1. Obtener la primera OT de la legalización
            String sqlOT = "SELECT ot FROM linealegalizacion WHERE legalizacion = (SELECT codigo FROM legalizacion WHERE codigolegalizacion = '"
                    + codigoLegalizacion + "') LIMIT 1";
            ResultSet rsOT = conexion.funcionConsultar(sqlOT);
            String[][] datosOT = ConexionDatos.armarArreglo(rsOT);

            if (datosOT != null && !datosOT[0][0].equals("0")) {
                String otId = datosOT[0][0];
                // 2. Consultar el cliente de esa OT
                String sqlCL = "SELECT cliente FROM otviaticos WHERE numero = '" + otId + "'";
                ResultSet rsCL = conexion.funcionConsultar(sqlCL);
                String[][] datosCL = ConexionDatos.armarArreglo(rsCL);

                if (datosCL != null && !datosCL[0][0].equals("0")) {
                    // 3. Capturar el nombre del Cliente
                    clienteEncontrado = datosCL[0][0];
                }
            }
        } catch (Exception e) {
            System.err.println("Error obteniendo el cliente de la OT: " + e.getMessage());
        }
        // Sobrescribir el campo de Obra con el cliente encontrado o el guion
        datosL[0][11] = clienteEncontrado;
        // ----------------------------------------

        // Configuración archivo (carpeta temp + nombre único)
        File carpetaTemp = new File("temp");
        if (!carpetaTemp.exists())
            carpetaTemp.mkdirs();

        dirArchivo = "temp" + File.separator + "Reporte_Viaticos_" + codigoLegalizacion + "_"
                + System.currentTimeMillis() + ".pdf";
        Document documento = new Document(PageSize.LETTER, 30, 30, 30, 30);
        PdfWriter writer = PdfWriter.getInstance(documento, new FileOutputStream(dirArchivo));
        writer.setPageEvent(this);

        documento.open();

        // 1. Encabezado Triple
        documento.add(crearEncabezadoPrincipal(codigoLegalizacion));
        documento.add(new Paragraph("\n", fontMini));

        // 2. Información del Empleado y Obra
        documento.add(crearTablaInfoEmpleado(datosL[0]));
        documento.add(new Paragraph("\n", fontMini));

        // 3. Tablas de Resumen y Totales (Lado a Lado usando una tabla contenedora)
        documento.add(crearBloqueCentral(codigoLegalizacion, datosL[0]));
        documento.add(new Paragraph("\n", fontMini));

        // 3.5 Matriz DÍA vs RUBRO
        documento.add(crearMatrizDiaRubro(codigoLegalizacion));
        documento.add(new Paragraph("\n", fontMini));

        // 4. Detalle de Gastos
        documento.add(crearTablaDetalleGastos(codigoLegalizacion));

        documento.close();

        // Abrir automáticamente
        abrirArchivo();
    }

    private PdfPTable crearEncabezadoPrincipal(String radicado) throws Exception {
        PdfPTable table = new PdfPTable(3);
        table.setWidthPercentage(100);
        table.setWidths(new float[] { 20, 60, 20 });

        // Logo
        PdfPCell cellLogo = new PdfPCell();
        if (imageLogo != null) {
            Image img = Image.getInstance(imageLogo, null);
            img.scaleToFit(80, 50);
            cellLogo.addElement(img);
        }
        cellLogo.setBorder(PdfPCell.BOX);
        cellLogo.setHorizontalAlignment(Element.ALIGN_CENTER);
        table.addCell(cellLogo);

        // Título
        PdfPCell cellTitulo = crearCelda("REPORTE DE GASTOS DE VIATICOS", fontTitulo, Element.ALIGN_CENTER, null, 10);
        cellTitulo.setVerticalAlignment(Element.ALIGN_MIDDLE);
        table.addCell(cellTitulo);

        // Radicado (Caja Azul)
        PdfPTable tableRad = new PdfPTable(1);
        tableRad.addCell(crearCelda("RADICADO", fontNegritaBlancas, Element.ALIGN_CENTER, AZUL_NAVY, 3));
        tableRad.addCell(crearCelda(radicado, fontNegrita, Element.ALIGN_CENTER, Color.WHITE, 5));

        PdfPCell cellRad = new PdfPCell(tableRad);
        cellRad.setPadding(5);
        table.addCell(cellRad);

        return table;
    }

    private PdfPTable crearTablaInfoEmpleado(String[] datos) throws Exception {
        String[] headers = { "FECHA ENTREGA", "NOMBRE", "CEDULA", "OBRA", "CIUDAD", "OT", "TOTAL SOLICITADO",
                "TOTAL APROBADO" };
        PdfPTable table = new PdfPTable(headers.length);
        table.setWidthPercentage(100);

        // Encabezados
        for (String h : headers) {
            table.addCell(crearCelda(h, fontNegritaBlancas, Element.ALIGN_CENTER, AZUL_NAVY, 3));
        }

        // Datos (Mapeo basado en el SELECT de legalización)
        table.addCell(crearCelda(datos[4], fontNormal, Element.ALIGN_CENTER, GRIS_CLARO, 4)); // fechaaplicacion
        table.addCell(crearCelda(datos[6], fontNormal, Element.ALIGN_CENTER, GRIS_CLARO, 4)); // nombreempleado
        table.addCell(crearCelda(datos[5], fontNormal, Element.ALIGN_CENTER, GRIS_CLARO, 4)); // empleado (como doc)
        table.addCell(crearCelda(datos[11], fontNormal, Element.ALIGN_CENTER, GRIS_CLARO, 4)); // obra-Cliente Primera
                                                                                               // OT
        table.addCell(crearCelda(datos[15], fontNormal, Element.ALIGN_CENTER, GRIS_CLARO, 4)); // ciudad
        table.addCell(crearCelda(datos[13], fontNormal, Element.ALIGN_CENTER, GRIS_CLARO, 4)); // centrocosto (OT

        table.addCell(crearCelda(formateadorMoneda.format(Double.parseDouble(datos[8])), fontNormal,
                Element.ALIGN_RIGHT, GRIS_CLARO, 4));
        table.addCell(crearCelda(formateadorMoneda.format(Double.parseDouble(datos[8])), fontNormal,
                Element.ALIGN_RIGHT, GRIS_CLARO, 4));

        return table;
    }

    private PdfPTable crearBloqueCentral(String id, String[] legalizacion) throws Exception {
        // 1. Obtención de datos reales
        String sqlOT = "SELECT ot, centrocosto, subcentrocosto, SUM(valorconfactura + valorsinfactura) FROM linealegalizacion WHERE legalizacion = (SELECT codigo FROM legalizacion WHERE codigolegalizacion = '"
                + id + "') GROUP BY ot, centrocosto, subcentrocosto";
        String[][] datosOT = ConexionDatos.armarArreglo(conexion.funcionConsultar(sqlOT));

        String sqlR = "SELECT categoria, SUM(valorconfactura + valorsinfactura) FROM linealegalizacion WHERE legalizacion = (SELECT codigo FROM legalizacion WHERE codigolegalizacion = '"
                + id + "') GROUP BY categoria";
        String[][] datosR = ConexionDatos.armarArreglo(conexion.funcionConsultar(sqlR));

        double totalGasto = Double.parseDouble(legalizacion[8]);

        // 2. Tabla de RESUMEN (7 columnas: OT, VALOR, CC, SUB, RUBRO, SOLICITADO,
        // APROBADO)
        float[] relativeWidths = { 10, 15, 10, 10, 25, 15, 15 };
        PdfPTable tableRes = new PdfPTable(relativeWidths);
        tableRes.setWidthPercentage(100);

        // Encabezados
        tableRes.addCell(crearCelda("OT", fontNegritaBlancas, Element.ALIGN_CENTER, AZUL_NAVY, 2));
        tableRes.addCell(crearCelda("CENTRO DE COSTO", fontNegritaBlancas, Element.ALIGN_CENTER, AZUL_NAVY, 2));
        tableRes.addCell(crearCelda("SUB CENTRO", fontNegritaBlancas, Element.ALIGN_CENTER, AZUL_NAVY, 2));
        tableRes.addCell(crearCelda("VALOR TOTAL APROBADO", fontNegritaBlancas, Element.ALIGN_CENTER, AZUL_NAVY, 2));
        tableRes.addCell(crearCelda("RUBRO", fontNegritaBlancas, Element.ALIGN_CENTER, AZUL_NAVY, 2));
        tableRes.addCell(crearCelda("TOTAL SOLICITADO", fontNegritaBlancas, Element.ALIGN_CENTER, AZUL_NAVY, 2));
        tableRes.addCell(crearCelda("TOTAL APROBADO", fontNegritaBlancas, Element.ALIGN_CENTER, AZUL_NAVY, 2));

        // Determinamos el número de filas real (sin filas en blanco extras)
        int nOT = (datosOT != null && !datosOT[0][0].equals("0")) ? datosOT.length : 0;
        int nR = (datosR != null && !datosR[0][0].equals("0")) ? datosR.length : 0;
        int rowsNeeded = Math.max(nOT, nR + 1); // +1 para la fila del total de rubros

        for (int i = 0; i < rowsNeeded; i++) {
            // -- Bloque OT (Columnas 1-4) --
            if (i < nOT) {
                tableRes.addCell(crearCelda(datosOT[i][0], fontNormal, Element.ALIGN_CENTER, null, 2));
                tableRes.addCell(crearCelda(datosOT[i][1], fontNormal, Element.ALIGN_CENTER, null, 2));
                tableRes.addCell(crearCelda(datosOT[i][2], fontNormal, Element.ALIGN_CENTER, null, 2));
                tableRes.addCell(crearCelda(formateadorMoneda.format(Double.parseDouble(datosOT[i][3])), fontNormal,
                        Element.ALIGN_RIGHT, null, 2));
            } else {
                for (int c = 0; c < 4; c++)
                    tableRes.addCell(crearCelda("", fontNormal, Element.ALIGN_CENTER, null, 2));
            }

            // -- Bloque Rubro (Columnas 5-7) --
            if (i < nR) {
                // Column 5: Rubro (con fondo gris)
                tableRes.addCell(crearCelda(datosR[i][0] + ":", fontNormal, Element.ALIGN_LEFT, GRIS_CLARO, 2));
                // Column 6: Solicitado (fondo blanco)
                tableRes.addCell(crearCelda(formateadorMoneda.format(Double.parseDouble(datosR[i][1])), fontNormal,
                        Element.ALIGN_RIGHT, null, 2));
                // Column 7: Aprobado (fondo blanco)
                tableRes.addCell(crearCelda(formateadorMoneda.format(Double.parseDouble(datosR[i][1])), fontNormal,
                        Element.ALIGN_RIGHT, null, 2));
            } else if (i == nR) {
                // Fila de Total
                tableRes.addCell(crearCelda("Total", fontNegrita, Element.ALIGN_LEFT, GRIS_CLARO, 2));
                tableRes.addCell(
                        crearCelda(formateadorMoneda.format(totalGasto), fontNegrita, Element.ALIGN_RIGHT, null, 2));
                tableRes.addCell(
                        crearCelda(formateadorMoneda.format(totalGasto), fontNegrita, Element.ALIGN_RIGHT, null, 2));
            } else {
                for (int c = 0; c < 3; c++)
                    tableRes.addCell(crearCelda("", fontNormal, Element.ALIGN_CENTER, null, 2));
            }
        }

        // Retornamos directamente la tabla de resumen
        return tableRes;
    }

    /**
     * Crea una tabla-matriz que cruza FECHAS (filas) contra RUBROS/CATEGORÍAS
     * (columnas)
     * mostrando valores totales por celda. Replica la sección 3.5 del reporte web.
     */
    private PdfPTable crearMatrizDiaRubro(String id) throws Exception {
        // Consulta agrupada por fecha y categoría
        String sql = "SELECT fecharealgasto, categoria, SUM(valorconfactura + valorsinfactura) "
                + "FROM linealegalizacion WHERE legalizacion = (SELECT codigo FROM legalizacion "
                + "WHERE codigolegalizacion = '" + id + "') GROUP BY fecharealgasto, categoria "
                + "ORDER BY fecharealgasto, categoria";
        ResultSet rs = conexion.funcionConsultar(sql);
        String[][] datos = ConexionDatos.armarArreglo(rs);

        // Estructuras: TreeMap ordena por fecha, TreeSet ordena categorías
        TreeMap<String, Map<String, Double>> dayRubroMap = new TreeMap<>();
        TreeSet<String> categoriasSet = new TreeSet<>();

        if (datos != null && !datos[0][0].equals("0")) {
            for (String[] fila : datos) {
                String fecha = fila[0];
                String categoria = fila[1];
                double valor = Double.parseDouble(fila[2]);

                categoriasSet.add(categoria);
                dayRubroMap.computeIfAbsent(fecha, k -> new TreeMap<>());
                dayRubroMap.get(fecha).merge(categoria, valor, Double::sum);
            }
        }

        // Lista ordenada de categorías para indexar columnas
        List<String> categorias = new ArrayList<>(categoriasSet);
        int numCategorias = categorias.size();

        // Número de columnas: 1 (FECHA) + N categorías + 1 (TOTAL)
        int numCols = numCategorias + 2;
        PdfPTable table = new PdfPTable(numCols);
        table.setWidthPercentage(100);

        // Anchos dinámicos: columna fecha fija, resto proporcional
        float[] widths = new float[numCols];
        widths[0] = 14; // Columna FECHA
        float restWidth = (100f - 14f - 14f) / Math.max(numCategorias, 1);
        for (int i = 1; i <= numCategorias; i++) {
            widths[i] = restWidth;
        }
        widths[numCols - 1] = 14; // Columna TOTAL
        table.setWidths(widths);

        // Fuente dinámica según cantidad de columnas
        Font fontMatriz;
        float padding;
        if (numCategorias > 8) {
            fontMatriz = FontFactory.getFont("Helvetica", 5, Font.NORMAL, Color.BLACK);
            padding = 1;
        } else if (numCategorias > 5) {
            fontMatriz = FontFactory.getFont("Helvetica", 6, Font.NORMAL, Color.BLACK);
            padding = 1.5f;
        } else {
            fontMatriz = fontNormal; // 8pt
            padding = 2;
        }
        Font fontMatrizBold = FontFactory.getFont("Helvetica",
                fontMatriz.getSize(), Font.BOLD, Color.BLACK);
        Font fontMatrizBoldWhite = FontFactory.getFont("Helvetica",
                fontMatriz.getSize(), Font.BOLD, Color.WHITE);
        Font fontMatrizHeader = FontFactory.getFont("Helvetica",
                fontMatriz.getSize(), Font.BOLD, Color.WHITE);

        // === ENCABEZADO ===
        table.addCell(crearCelda("FECHA /\nRUBRO", fontMatrizHeader, Element.ALIGN_CENTER, AZUL_NAVY, padding));
        for (String cat : categorias) {
            table.addCell(crearCelda(cat.toUpperCase(), fontMatrizHeader, Element.ALIGN_CENTER, AZUL_NAVY, padding));
        }
        table.addCell(crearCelda("TOTAL", fontMatrizHeader, Element.ALIGN_CENTER, AZUL_NAVY, padding));

        // === CUERPO: Una fila por fecha ===
        double[] totalPorCategoria = new double[numCategorias];
        double granTotal = 0;

        int rowIdx = 0;
        for (Map.Entry<String, Map<String, Double>> entry : dayRubroMap.entrySet()) {
            String fecha = entry.getKey();
            Map<String, Double> catMap = entry.getValue();
            Color bg = (rowIdx % 2 == 0) ? Color.WHITE : GRIS_CLARO;

            // Columna Fecha (fondo gris siempre para resaltar)
            table.addCell(crearCelda(fecha, fontMatriz, Element.ALIGN_CENTER,
                    new Color(220, 220, 220), padding));

            double totalDia = 0;
            for (int c = 0; c < numCategorias; c++) {
                double valor = catMap.getOrDefault(categorias.get(c), 0.0);
                totalDia += valor;
                totalPorCategoria[c] += valor;

                // Mostrar "-" si el valor es 0 para reducir ruido visual
                String texto = valor > 0 ? formateadorMoneda.format(valor) : "-";
                table.addCell(crearCelda(texto, fontMatriz, Element.ALIGN_RIGHT, bg, padding));
            }

            // Columna TOTAL del día (fondo gris, negrita)
            granTotal += totalDia;
            table.addCell(crearCelda(formateadorMoneda.format(totalDia), fontMatrizBold,
                    Element.ALIGN_RIGHT, new Color(220, 220, 220), padding));

            rowIdx++;
        }

        // === FILA TOTALES (Fondo azul navy, texto blanco en TODA la fila) ===
        table.addCell(crearCelda("TOTAL", fontMatrizBoldWhite, Element.ALIGN_CENTER, AZUL_NAVY, padding));
        for (int c = 0; c < numCategorias; c++) {
            table.addCell(crearCelda(formateadorMoneda.format(totalPorCategoria[c]),
                    fontMatrizBoldWhite, Element.ALIGN_RIGHT, AZUL_NAVY, padding));
        }
        table.addCell(crearCelda(formateadorMoneda.format(granTotal),
                fontMatrizBoldWhite, Element.ALIGN_RIGHT, AZUL_NAVY, padding));

        return table;
    }

    private PdfPTable crearTablaDetalleGastos(String id) throws Exception {
        String sql = "SELECT * FROM linealegalizacion WHERE legalizacion = (SELECT codigo FROM legalizacion WHERE codigolegalizacion = '"
                + id + "') ORDER BY fecharealgasto ASC";
        ResultSet rs = conexion.funcionConsultar(sql);
        String[][] datos = ConexionDatos.armarArreglo(rs);

        String[] headers = { "FECHA REAL", "DESCRIPCION", "OT", "C. COSTO", "VALOR CON FACT", "VALOR SIN FACT",
                "TOTAL" };
        PdfPTable table = new PdfPTable(headers.length);
        table.setWidthPercentage(100);
        table.setWidths(new float[] { 12, 35, 8, 10, 12, 12, 11 });

        for (String h : headers)
            table.addCell(crearCelda(h, fontNegritaBlancas, Element.ALIGN_CENTER, AZUL_NAVY, 3));

        if (datos != null && !datos[0][0].equals("0")) {
            double sumCon = 0, sumSin = 0, sumTotal = 0;
            for (int i = 0; i < datos.length; i++) {
                Color bg = (i % 2 == 0) ? Color.WHITE : GRIS_CLARO;
                double vCon = Double.parseDouble(datos[i][8]);
                double vSin = Double.parseDouble(datos[i][9]);
                double total = vCon + vSin;

                sumCon += vCon;
                sumSin += vSin;
                sumTotal += total;

                table.addCell(crearCelda(datos[i][3], fontNormal, Element.ALIGN_CENTER, bg, 3));
                table.addCell(crearCelda(datos[i][4] + ": " + datos[i][10], fontNormal, Element.ALIGN_LEFT, bg, 3));
                table.addCell(crearCelda(datos[i][5], fontNormal, Element.ALIGN_CENTER, bg, 3));
                table.addCell(crearCelda(datos[i][6], fontNormal, Element.ALIGN_CENTER, bg, 3));
                table.addCell(crearCelda(formateadorMoneda.format(vCon), fontNormal, Element.ALIGN_RIGHT, bg, 3));
                table.addCell(crearCelda(formateadorMoneda.format(vSin), fontNormal, Element.ALIGN_RIGHT, bg, 3));
                table.addCell(crearCelda(formateadorMoneda.format(total), fontNegrita, Element.ALIGN_RIGHT, bg, 3));
            }
            // Fila de Totales
            table.addCell(crearCelda("TOTALES", fontNegritaBlancas, Element.ALIGN_CENTER, AZUL_NAVY, 3, 4));
            table.addCell(crearCelda(formateadorMoneda.format(sumCon), fontNegritaBlancas, Element.ALIGN_RIGHT,
                    AZUL_NAVY, 3));
            table.addCell(crearCelda(formateadorMoneda.format(sumSin), fontNegritaBlancas, Element.ALIGN_RIGHT,
                    AZUL_NAVY, 3));
            table.addCell(crearCelda(formateadorMoneda.format(sumTotal), fontNegritaBlancas, Element.ALIGN_RIGHT,
                    AZUL_NAVY, 3));
        }

        return table;
    }

    private void abrirArchivo() {
        try {
            File file = new File(dirArchivo);
            if (file.exists()) {
                Runtime.getRuntime().exec("cmd /c \"" + dirArchivo + "\"");
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
