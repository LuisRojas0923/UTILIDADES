import java.io.FileOutputStream;
import java.io.File;
import java.awt.Color;
import java.awt.Toolkit;
import java.sql.ResultSet;
import javax.swing.JOptionPane;
import com.lowagie.text.*;
import com.lowagie.text.pdf.*;

import com.lowagie.text.Document;
import com.lowagie.text.DocumentException;
import com.lowagie.text.Paragraph;
import com.lowagie.text.pdf.PdfWriter;
import com.lowagie.text.pdf.*;
import com.lowagie.text.*;
import java.sql.*;
import java.awt.image.BufferedImage;
import java.awt.*;

/**
 * Generador de Reporte de Gastos de Viáticos siguiendo el diseño Refridcol.
 */
public class ReporteGastosViaticos extends BaseInforme {

    private ConexionDatos conexion;
    private String dirArchivo;
    private String logoPath = null;
    private byte[] logoBytes = null;
    private com.lowagie.text.Image imgLogo = null;
    private com.lowagie.text.Image imgMarcaAgua = null;

    public ReporteGastosViaticos(ConexionDatos inConexion) {
        super();
        this.conexion = inConexion;
        cargarLogo();
    }

    private void cargarLogo() {
        try {
            String pathLocal = "\\images\\logoBarra1.png";
            File archivoLogo = new File(pathLocal);
            if (archivoLogo.exists()) {
                logoPath = archivoLogo.getAbsolutePath();
            } else {
                // Imagen tmpImagen = new Imagen(null, conexion);
                // tmpImagen.cargar("codigo", "3");
                // logoBytes = (byte[]) tmpImagen.representacion;
            }
        } catch (Exception e) {
        }
    }

    private void inicializarLogo() {
        if (imgLogo != null)
            return;
        try {
            Imagen tmpImagen = new Imagen(null, conexion);
            tmpImagen.cargar("codigo", "1");

            java.awt.Image imagenMembrete = null;
            try {
                imagenMembrete = ImageConverter.convertToImage(tmpImagen.representacion);
            } catch (Exception e) {
                BufferedImage imagen = new BufferedImage(200, 200, BufferedImage.TYPE_INT_ARGB);
                Graphics2D g2d = imagen.createGraphics();
                g2d.setColor(new java.awt.Color(0, 0, 0, 0));
                g2d.fillRect(0, 0, imagen.getWidth(), imagen.getHeight());
                imagenMembrete = (java.awt.Image) imagen;
            }
            imgLogo = com.lowagie.text.Image.getInstance(imagenMembrete, null);
            imgLogo.scaleToFit(60, 40); // Tamaño para encabezado
            imgLogo.setAlignment(com.lowagie.text.Image.DEFAULT);

            // Cargar Marca de Agua desde BD (Código 4)
            try {
                Imagen tmpImagenMarca = new Imagen(null, conexion);
                tmpImagenMarca.cargar("codigo", "4"); // Código sugerido por usuario

                if (tmpImagenMarca.representacion != null) {
                    java.awt.Image awtMarca = ImageConverter.convertToImage(tmpImagenMarca.representacion);
                    imgMarcaAgua = com.lowagie.text.Image.getInstance(awtMarca, null);
                }
            } catch (Exception e) {
                System.out.println("No se pudo cargar la marca de agua (DB Cod 4): " + e.getMessage());
            }

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    @Override
    public void generar(String codigoLegalizacion) throws Exception {
        inicializarLogo(); // Cargar logo antes de empezar
        // Consultar cabecera
        String sqlL = "SELECT * FROM legalizacion WHERE codigolegalizacion = '" + codigoLegalizacion + "'";
        ResultSet rsL = conexion.funcionConsultar(sqlL);
        // Usamos armarArreglo heredado de BaseInforme, no de ConexionDatos
        String[][] datosL = armarArreglo(rsL);

        if (datosL == null || datosL[0].length < 2) {
            throw new Exception("No se encontró la legalización: " + codigoLegalizacion);
        }

        // Configuración archivo
        File carpetaTemp = new File("temp");
        if (!carpetaTemp.exists()) {
            carpetaTemp.mkdirs();
        }
        dirArchivo = "temp" + File.separator + "Reporte_Viaticos_" + codigoLegalizacion + " "
                + FileReaderWriter.generarNombreArchivoTemporal() + ".pdf";
        Document documento = new Document(PageSize.LETTER, 30, 30, 30, 30);
        PdfWriter writer = PdfWriter.getInstance(documento, new FileOutputStream(dirArchivo));
        writer.setPageEvent(this);

        documento.open();

        // ============================================================================
        // SECCIÓN 1: ENCABEZADO TRIPLE (Logo, Título, Radicado)
        // ============================================================================
        documento.add(crearEncabezadoPrincipal(codigoLegalizacion));
        documento.add(new Paragraph("\n", fontMini));

        // ============================================================================
        // SECCIÓN 2: INFORMACIÓN DEL EMPLEADO
        // ============================================================================
        documento.add(crearTablaInfoEmpleado(datosL[0]));
        documento.add(new Paragraph("\n", fontMini));

        // ============================================================================
        // SECCIÓN 3: BLOQUE CENTRAL (RESUMEN POR OT / CENTRO COSTO / RUBRO)
        // ============================================================================
        documento.add(crearBloqueCentral(codigoLegalizacion, datosL[0]));
        documento.add(new Paragraph("\n", fontMini));

        // ============================================================================
        // SECCIÓN 4: DETALLE DE GASTOS (TABLA INFERIOR)
        // ============================================================================
        documento.add(crearTablaDetalleGastos(codigoLegalizacion));

        documento.close();

        // Abrir automáticamente
        abrirArchivo();
    }

    // --------------------------------------------------------------------------------
    // MÉTODOS DE CONSTRUCCIÓN DE BLOQUES
    // --------------------------------------------------------------------------------

    private PdfPTable crearEncabezadoPrincipal(String radicado) throws Exception {
        PdfPTable table = new PdfPTable(3);
        int[] anchosTablaDetalles = { 30, 50, 20 };
        table.setWidths(anchosTablaDetalles);
        table.setWidthPercentage(100);
        // table.setWidths(new float[] { 20, 60, 20 });

        // table.setWidths(new float[] { 20, 60, 20 });

        // Logo (Usando campo inicializado)
        if (imgLogo == null)
            inicializarLogo();

        // Logo
        PdfPCell cellLogo = new PdfPCell();
        if (imgLogo != null) {
            cellLogo.addElement(imgLogo);
        }
        if (logoPath != null || logoBytes != null) {
            // Image img = (logoPath != null) ? Image.getInstance(logoPath) :
            // Image.getInstance(logoBytes);
            // img.scaleToFit(40, 40); // Ajustado manualmente por usuario

        }
        cellLogo.setBorder(PdfPCell.BOX);
        // cellLogo.setPadding(8);
        cellLogo.setFixedHeight(50f); // Ajustado a 50pt
        cellLogo.setHorizontalAlignment(Element.ALIGN_CENTER);
        cellLogo.setVerticalAlignment(Element.ALIGN_MIDDLE);
        table.addCell(cellLogo);

        // Título
        PdfPCell cellTitulo = crearCelda("REPORTE DE GASTOS DE VIATICOS", fontTitulo, Element.ALIGN_CENTER, null, 2);
        cellTitulo.setVerticalAlignment(Element.ALIGN_MIDDLE);
        cellTitulo.setFixedHeight(50f); // Ajustado a 50pt
        table.addCell(cellTitulo);

        // Radicado (Caja Azul)
        PdfPTable tableRad = new PdfPTable(1);

        tableRad.addCell(crearCelda("RADICADO", fontNegritaBlancas, Element.ALIGN_CENTER, AZUL_NAVY, 1));
        tableRad.addCell(crearCelda(radicado, fontNegrita, Element.ALIGN_CENTER, Color.WHITE, 2));

        PdfPCell cellRad = new PdfPCell(tableRad);
        cellRad.setPadding(10); // Ajustado manualmente
        cellRad.setFixedHeight(50f); // Ajustado a 50pt
        table.addCell(cellRad);

        return table;
    }

    private PdfPTable crearTablaInfoEmpleado(String[] datos) throws Exception {
        String[] headers = { "FECHA ENTREGA", "NOMBRE", "CEDULA", "TOTAL SOLICITADO", "TOTAL APROBADO" };
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

        table.addCell(crearCelda(formateadorMoneda.format(Double.parseDouble(datos[8])), fontNormal,
                Element.ALIGN_RIGHT, GRIS_CLARO, 4));
        table.addCell(crearCelda(formateadorMoneda.format(Double.parseDouble(datos[8])), fontNormal,
                Element.ALIGN_RIGHT, GRIS_CLARO, 4));

        return table;
    }

    private PdfPTable crearBloqueCentral(String id, String[] legalizacion) throws Exception {
        // 1. Obtención de datos reales
        // 1a. Agrupamiento por OT
        String sqlOT = "SELECT ot, SUM(valorconfactura + valorsinfactura) FROM linealegalizacion WHERE legalizacion = (SELECT codigo FROM legalizacion WHERE codigolegalizacion = '"
                + id + "') GROUP BY ot";
        String[][] datosOT = ConexionDatos.armarArreglo(conexion.funcionConsultar(sqlOT));

        // 1b. Agrupamiento por Centro de Costo (Incluyendo subcentro)
        String sqlCC = "SELECT centrocosto, subcentrocosto, SUM(valorconfactura + valorsinfactura) FROM linealegalizacion WHERE legalizacion = (SELECT codigo FROM legalizacion WHERE codigolegalizacion = '"
                + id + "') GROUP BY centrocosto, subcentrocosto";
        String[][] datosCC = ConexionDatos.armarArreglo(conexion.funcionConsultar(sqlCC));

        // 1c. Agrupamiento por Rubro
        String sqlR = "SELECT categoria, SUM(valorconfactura + valorsinfactura) FROM linealegalizacion WHERE legalizacion = (SELECT codigo FROM legalizacion WHERE codigolegalizacion = '"
                + id + "') GROUP BY categoria";
        String[][] datosR = ConexionDatos.armarArreglo(conexion.funcionConsultar(sqlR));

        double totalGasto = Double.parseDouble(legalizacion[8]);

        // 2. Tabla de RESUMEN (9 columnas: OT|Valor | CC|Valor | Rubro|Sol|Apr)
        float[] relativeWidths = { 10, 10, 10, 10, 20, 10, 10 };
        PdfPTable tableRes = new PdfPTable(relativeWidths);
        tableRes.setWidthPercentage(100);

        // Encabezados
        tableRes.addCell(crearCelda("OT", fontNegritaBlancas, Element.ALIGN_CENTER, AZUL_NAVY, 2));
        tableRes.addCell(crearCelda("VALOR TOTAL", fontNegritaBlancas, Element.ALIGN_CENTER, AZUL_NAVY, 2));

        tableRes.addCell(crearCelda("C.COSTO", fontNegritaBlancas, Element.ALIGN_CENTER, AZUL_NAVY, 2));
        tableRes.addCell(crearCelda("VALOR TOTAL", fontNegritaBlancas, Element.ALIGN_CENTER, AZUL_NAVY, 2));

        tableRes.addCell(crearCelda("RUBRO", fontNegritaBlancas, Element.ALIGN_CENTER, AZUL_NAVY, 2));
        tableRes.addCell(crearCelda("SOLICITADO", fontNegritaBlancas, Element.ALIGN_CENTER, AZUL_NAVY, 2));
        tableRes.addCell(crearCelda("APROBADO", fontNegritaBlancas, Element.ALIGN_CENTER, AZUL_NAVY, 2));

        // Determinamos el número de filas real
        int nOT = (datosOT != null && !datosOT[0][0].equals("0")) ? datosOT.length : 0;
        int nCC = (datosCC != null && !datosCC[0][0].equals("0")) ? datosCC.length : 0;
        int nR = (datosR != null && !datosR[0][0].equals("0")) ? datosR.length : 0;

        // La altura del cuerpo es el máximo de filas de datos
        int rowsBody = Math.max(nOT, Math.max(nCC, nR));

        for (int i = 0; i < rowsBody; i++) {
            // -- Bloque OT (Columns 1-2) --
            if (i < nOT) {
                // OT Code con Fondo Gris
                tableRes.addCell(crearCelda(datosOT[i][0], fontNormal, Element.ALIGN_CENTER, GRIS_CLARO, 2));
                tableRes.addCell(crearCelda(formateadorMoneda.format(Double.parseDouble(datosOT[i][1])), fontNormal,
                        Element.ALIGN_RIGHT, null, 2));
            } else {
                tableRes.addCell(crearCelda("", fontNormal, Element.ALIGN_CENTER, null, 2));
                tableRes.addCell(crearCelda("", fontNormal, Element.ALIGN_CENTER, null, 2));
            }

            // -- Bloque CC (Columns 3-4) --
            if (i < nCC) {
                // CC Code con Fondo Gris (Concatenado con Subcentro)
                String ccDisplay = datosCC[i][0];
                if (datosCC[i][1] != null && !datosCC[i][1].isEmpty()) {
                    ccDisplay += "-" + datosCC[i][1];
                }
                tableRes.addCell(crearCelda(ccDisplay, fontNormal, Element.ALIGN_CENTER, GRIS_CLARO, 2));
                // El valor está ahora en el índice 2 debido a la columna extra en el SELECT
                tableRes.addCell(crearCelda(formateadorMoneda.format(Double.parseDouble(datosCC[i][2])), fontNormal,
                        Element.ALIGN_RIGHT, null, 2));
            } else {
                tableRes.addCell(crearCelda("", fontNormal, Element.ALIGN_CENTER, null, 2));
                tableRes.addCell(crearCelda("", fontNormal, Element.ALIGN_CENTER, null, 2));
            }

            // -- Bloque Rubro (Columns 5-7) --
            if (i < nR) {
                // Rubro Name con Fondo Gris
                tableRes.addCell(crearCelda(datosR[i][0] + ":", fontNormal, Element.ALIGN_LEFT, GRIS_CLARO, 2));
                tableRes.addCell(crearCelda(formateadorMoneda.format(Double.parseDouble(datosR[i][1])), fontNormal,
                        Element.ALIGN_RIGHT, null, 2));
                tableRes.addCell(crearCelda(formateadorMoneda.format(Double.parseDouble(datosR[i][1])), fontNormal,
                        Element.ALIGN_RIGHT, null, 2));
            } else {
                tableRes.addCell(crearCelda("", fontNormal, Element.ALIGN_CENTER, null, 2));
                tableRes.addCell(crearCelda("", fontNormal, Element.ALIGN_CENTER, null, 2));
                tableRes.addCell(crearCelda("", fontNormal, Element.ALIGN_CENTER, null, 2));
            }
        }

        // --- FILA DE TOTALES (Siempre al final, alineada) ---

        // Total OT
        tableRes.addCell(crearCelda("TOTAL", fontNegrita, Element.ALIGN_CENTER, GRIS_CLARO, 2));
        tableRes.addCell(crearCelda(formateadorMoneda.format(totalGasto), fontNegrita, Element.ALIGN_RIGHT, null, 2));

        // Total CC
        tableRes.addCell(crearCelda("TOTAL", fontNegrita, Element.ALIGN_CENTER, GRIS_CLARO, 2));
        tableRes.addCell(crearCelda(formateadorMoneda.format(totalGasto), fontNegrita, Element.ALIGN_RIGHT, null, 2));

        // Total Rubro
        tableRes.addCell(crearCelda("TOTAL", fontNegrita, Element.ALIGN_LEFT, GRIS_CLARO, 2));
        tableRes.addCell(crearCelda(formateadorMoneda.format(totalGasto), fontNegrita, Element.ALIGN_RIGHT, null, 2));
        tableRes.addCell(crearCelda(formateadorMoneda.format(totalGasto), fontNegrita, Element.ALIGN_RIGHT, null, 2));

        // Retornamos directamente la tabla de resumen
        return tableRes;
    }

    private PdfPTable crearTablaDetalleGastos(String id) throws Exception {
        String sql = "SELECT * FROM linealegalizacion WHERE legalizacion = (SELECT codigo FROM legalizacion WHERE codigolegalizacion = '"
                + id + "') ORDER BY fecharealgasto ASC";
        ResultSet rs = conexion.funcionConsultar(sql);
        String[][] datos = armarArreglo(rs);

        String[] headers = { "FECHA REAL", "DESCRIPCION", "OT", "C. COSTO", "S.C. COSTO", "VALOR CON FACTURA",
                "VALOR SIN FACTURA",
                "TOTAL" };
        PdfPTable table = new PdfPTable(headers.length);
        table.setWidthPercentage(100);
        table.setWidths(new float[] { 10, 30, 8, 8, 8, 11, 11, 10 }); // Ajuste de anchos para nueva columna

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

                table.addCell(crearCelda(datos[i][3], fontNormal, Element.ALIGN_CENTER, bg, 3)); // Fecha
                table.addCell(crearCelda(datos[i][4] + ": " + datos[i][10], fontNormal, Element.ALIGN_LEFT, bg, 3)); // Descripcion
                table.addCell(crearCelda(datos[i][5], fontNormal, Element.ALIGN_CENTER, bg, 3)); // OT
                table.addCell(crearCelda(datos[i][6], fontNormal, Element.ALIGN_CENTER, bg, 3)); // CC
                table.addCell(crearCelda(datos[i][7], fontNormal, Element.ALIGN_CENTER, bg, 3)); // Subcentro (Index 7)
                table.addCell(crearCelda(formateadorMoneda.format(vCon), fontNormal, Element.ALIGN_RIGHT, bg, 3));
                table.addCell(crearCelda(formateadorMoneda.format(vSin), fontNormal, Element.ALIGN_RIGHT, bg, 3));
                table.addCell(crearCelda(formateadorMoneda.format(total), fontNegrita, Element.ALIGN_RIGHT, bg, 3));
            }
            // Fila de Totales
            table.addCell(crearCelda("TOTALES", fontNegritaBlancas, Element.ALIGN_CENTER, AZUL_NAVY, 3, 5));
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

    @Override
    public void onEndPage(PdfWriter writer, Document document) {
        try {
            PdfContentByte cb = writer.getDirectContentUnder();

            // Dibujar marca de agua si existe
            if (imgMarcaAgua != null) {
                cb.saveState();

                // Transparencia al 100% (imagen ya trae opacidad)
                PdfGState gState = new PdfGState();
                gState.setFillOpacity(1.0f);
                gState.setStrokeOpacity(1.0f);
                cb.setGState(gState);

                // Posición: pie de página, esquina derecha
                float width = 80;
                float height = 60;
                float x = document.getPageSize().getWidth() - document.rightMargin() - width;
                float y = 15; // Cerca del borde inferior

                com.lowagie.text.Image imgFooter = com.lowagie.text.Image.getInstance(imgMarcaAgua);
                imgFooter.setAbsolutePosition(x, y);
                imgFooter.scaleToFit(width, height);

                cb.addImage(imgFooter);
                cb.restoreState();
            }

        } catch (Exception e) {
            System.out.println("Error dibujando marca de agua: " + e.getMessage());
        }
    }
}
