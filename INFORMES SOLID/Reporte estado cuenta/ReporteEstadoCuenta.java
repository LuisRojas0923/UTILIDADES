import java.io.FileOutputStream;
import java.io.File;
import java.awt.Color;
import java.awt.Image;
import java.sql.ResultSet;

import com.lowagie.text.*;
import com.lowagie.text.pdf.*;

/**
 * Generador de Reporte de Estado de Cuenta de Viáticos.
 */
public class ReporteEstadoCuenta extends BaseInforme {

    private ConexionDatos conexion;
    private String dirArchivo;
    private com.lowagie.text.Image imgLogo = null;
    private com.lowagie.text.Image imgMarcaAgua = null;

    public ReporteEstadoCuenta(ConexionDatos inConexion) {
        super();
        this.conexion = inConexion;
    }

    private void inicializarImagenes() {
        if (imgLogo != null)
            return;
        try {
            // Cargar Logo (Código 1)
            try {
                Imagen tmpImagen = new Imagen(null, conexion);
                tmpImagen.cargar("codigo", "1");
                if (tmpImagen.representacion != null) {
                    java.awt.Image awtLogo = ImageConverter.convertToImage(tmpImagen.representacion);
                    imgLogo = com.lowagie.text.Image.getInstance(awtLogo, null);
                    imgLogo.scaleToFit(60, 40);
                }
            } catch (Exception e) {
            }

            // Cargar Marca de Agua (Código 4)
            try {
                Imagen tmpImagenMarca = new Imagen(null, conexion);
                tmpImagenMarca.cargar("codigo", "4");
                if (tmpImagenMarca.representacion != null) {
                    java.awt.Image awtMarca = ImageConverter.convertToImage(tmpImagenMarca.representacion);
                    imgMarcaAgua = com.lowagie.text.Image.getInstance(awtMarca, null);
                }
            } catch (Exception e) {
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    @Override
    public void generar(String cedula) throws Exception {
        inicializarImagenes();

        // SQL adaptado de EstadoDeCuentaV2.sql pero filtrando solo por cedula
        String sql = "WITH movimientos AS (" +
                "    SELECT codigo, fechaaplicacion, empleado, nombreempleado, codigoconsignacion AS radicado, valor::numeric AS valor, 'CONSIGNACION' AS tipo, UPPER(TRIM(estado)) AS estado_limpio, observaciones FROM consignacion "
                +
                "    WHERE UPPER(TRIM(estado)) <> 'ANULADO' " +
                "    UNION ALL " +
                "    SELECT codigo, fechaaplicacion, empleado, nombreempleado, codigolegalizacion AS radicado, valortotal::numeric AS valor, 'LEGALIZACION' AS tipo, UPPER(TRIM(estado)) AS estado_limpio, observaciones FROM legalizacion "
                +
                "    WHERE UPPER(TRIM(estado)) <> 'ANULADO' " +
                ") " +
                "SELECT " +
                "    m.codigo AS \"CODIGO\", " +
                "    TO_CHAR(m.fechaaplicacion, 'DD/MM/YYYY') AS \"FECHA\", " +
                "    m.empleado AS \"CEDULA\", " +
                "    m.nombreempleado AS \"EMPLEADO\", " +
                "    m.radicado AS \"RADICADO\", " +
                "    CASE WHEN m.tipo = 'CONSIGNACION' AND m.estado_limpio = 'CONTABILIZADO' THEN m.valor ELSE 0 END AS \"C_CONT\", "
                +
                "    CASE WHEN m.tipo = 'LEGALIZACION' AND m.estado_limpio = 'CONTABILIZADO' THEN m.valor ELSE 0 END AS \"L_CONT\", "
                +
                "    CASE WHEN m.tipo = 'CONSIGNACION' AND m.estado_limpio = 'EN FIRME' THEN m.valor ELSE 0 END AS \"C_FIRM\", "
                +
                "    CASE WHEN m.tipo = 'LEGALIZACION' AND m.estado_limpio = 'EN FIRME' THEN m.valor ELSE 0 END AS \"L_FIRM\", "
                +
                "    CASE WHEN m.tipo = 'CONSIGNACION' AND m.estado_limpio = 'PENDIENTE' THEN m.valor ELSE 0 END AS \"C_PEND\", "
                +
                "    CASE WHEN m.tipo = 'LEGALIZACION' AND m.estado_limpio = 'PENDIENTE' THEN m.valor ELSE 0 END AS \"L_PEND\", "
                +
                "    SUM(CASE WHEN m.tipo = 'CONSIGNACION' THEN m.valor WHEN m.tipo = 'LEGALIZACION' THEN -m.valor ELSE 0 END) OVER (PARTITION BY m.empleado ORDER BY m.fechaaplicacion::timestamp ASC, m.codigo ASC) AS \"SALDO\", "
                +
                "    m.observaciones AS \"OBSERVACIONES\" " +
                "FROM movimientos m " +
                "WHERE m.empleado = '" + cedula + "' " +
                "ORDER BY m.fechaaplicacion::timestamp ASC, m.codigo ASC";

        ResultSet rs = conexion.funcionConsultar(sql);
        String[][] datos = armarArreglo(rs);

        if (datos == null || datos[0][0].equals("0")) {
            throw new Exception("No se encontraron movimientos para la cédula: " + cedula);
        }

        // Configuración archivo
        File carpetaTemp = new File("temp");
        if (!carpetaTemp.exists())
            carpetaTemp.mkdirs();

        dirArchivo = "temp" + File.separator + "EstadoCuenta_" + cedula + "_"
                + FileReaderWriter.generarNombreArchivoTemporal() + ".pdf";
        Document documento = new Document(PageSize.LETTER.rotate(), 30, 30, 30, 30); // Horizontal para mostrar todos
                                                                                     // los estados
        PdfWriter writer = PdfWriter.getInstance(documento, new FileOutputStream(dirArchivo));
        writer.setPageEvent(this);

        documento.open();

        // Encabezado Principal
        documento.add(crearEncabezadoPrincipal(cedula));
        documento.add(new Paragraph("\n", fontMini));

        // Información Empleado
        documento.add(crearTablaInfoEmpleado(datos[0]));
        documento.add(new Paragraph("\n", fontMini));

        // Tabla de Movimientos
        documento.add(crearTablaMovimientos(datos));

        documento.close();
        abrirArchivo();
    }

    private PdfPTable crearEncabezadoPrincipal(String cedula) throws Exception {
        PdfPTable table = new PdfPTable(3);
        table.setWidths(new int[] { 20, 60, 20 });
        table.setWidthPercentage(100);

        PdfPCell cellLogo = new PdfPCell();
        if (imgLogo != null)
            cellLogo.addElement(imgLogo);
        cellLogo.setBorder(PdfPCell.BOX);
        cellLogo.setFixedHeight(50f);
        cellLogo.setHorizontalAlignment(Element.ALIGN_CENTER);
        cellLogo.setVerticalAlignment(Element.ALIGN_MIDDLE);
        table.addCell(cellLogo);

        PdfPCell cellTitulo = crearCelda("ESTADO DE CUENTA DE VIATICOS", fontTitulo, Element.ALIGN_CENTER, null, 2);
        cellTitulo.setVerticalAlignment(Element.ALIGN_MIDDLE);
        cellTitulo.setFixedHeight(50f);
        table.addCell(cellTitulo);

        PdfPTable tableRad = new PdfPTable(1);
        tableRad.addCell(crearCelda("CEDULA VIATICANTE", fontNegritaBlancas, Element.ALIGN_CENTER, AZUL_NAVY, 2));
        tableRad.addCell(crearCelda(cedula, fontSubtitulo, Element.ALIGN_CENTER, Color.WHITE, 4));
        PdfPCell cellRad = new PdfPCell(tableRad);
        cellRad.setFixedHeight(50f);
        table.addCell(cellRad);

        return table;
    }

    private PdfPTable crearTablaInfoEmpleado(String[] primerFila) throws Exception {
        PdfPTable table = new PdfPTable(2);
        table.setWidthPercentage(100);
        table.setWidths(new int[] { 20, 80 });

        table.addCell(crearCelda("NOMBRE EMPLEADO", fontNegritaBlancas, Element.ALIGN_LEFT, AZUL_NAVY, 4));
        table.addCell(crearCelda(primerFila[3], fontSubtitulo, Element.ALIGN_LEFT, GRIS_CLARO, 4));

        return table;
    }

    private PdfPTable crearTablaMovimientos(String[][] datos) throws Exception {
        // Estructura de 10 columnas
        float[] relativeWidths = { 9, 9, 11, 11, 11, 11, 11, 11, 11, 24 };
        PdfPTable table = new PdfPTable(relativeWidths);
        table.setWidthPercentage(100);

        // --- FILA 1: ENCABEZADOS AGRUPADORES ---
        table.addCell(crearCelda("FECHA", fontNegritaBlancas, Element.ALIGN_CENTER, AZUL_NAVY, 5, 1));
        table.addCell(crearCelda("RADICADO", fontNegritaBlancas, Element.ALIGN_CENTER, AZUL_NAVY, 5, 1));
        table.addCell(crearCelda("CONTABILIZADO", fontNegritaBlancas, Element.ALIGN_CENTER, AZUL_NAVY, 5, 2));
        table.addCell(crearCelda("EN CANJE", fontNegritaBlancas, Element.ALIGN_CENTER, AZUL_NAVY, 5, 2));
        table.addCell(crearCelda("PENDIENTES", fontNegritaBlancas, Element.ALIGN_CENTER, AZUL_NAVY, 5, 2));
        table.addCell(crearCelda("SALDO", fontNegritaBlancas, Element.ALIGN_CENTER, AZUL_NAVY, 5, 1));
        table.addCell(crearCelda("OBSERVACIONES", fontNegritaBlancas, Element.ALIGN_CENTER, AZUL_NAVY, 5, 1));

        // --- FILA 2: SUB-ENCABEZADOS ---
        table.addCell(crearCelda("", fontNegritaBlancas, Element.ALIGN_CENTER, AZUL_NAVY, 2)); // Espacio para FECHA
        table.addCell(crearCelda("", fontNegritaBlancas, Element.ALIGN_CENTER, AZUL_NAVY, 2)); // Espacio para RADICADO

        // Consignación / Legalización repetido
        for (int i = 0; i < 3; i++) {
            table.addCell(crearCelda("CONSIGNACION", fontNegritaBlancas, Element.ALIGN_CENTER, AZUL_NAVY, 2));
            table.addCell(crearCelda("LEGALIZACION", fontNegritaBlancas, Element.ALIGN_CENTER, AZUL_NAVY, 2));
        }

        table.addCell(crearCelda("", fontNegritaBlancas, Element.ALIGN_CENTER, AZUL_NAVY, 2)); // Espacio para SALDO
        table.addCell(crearCelda("", fontNegritaBlancas, Element.ALIGN_CENTER, AZUL_NAVY, 2)); // Espacio para
                                                                                               // OBSERVACIONES

        // --- DATOS ---
        for (int i = 0; i < datos.length; i++) {
            Color bg = (i % 2 == 0) ? Color.WHITE : GRIS_CLARO;

            // Usamos fontNormal (8 pt) en lugar de fontMini (7 pt) para mejor lectura
            table.addCell(crearCelda(datos[i][1], fontNormal, Element.ALIGN_CENTER, bg, 3)); // FECHA
            table.addCell(crearCelda(datos[i][4], fontNormal, Element.ALIGN_LEFT, bg, 3)); // RADICADO

            // Valores numéricos - Si es 0 lo dejamos en blanco para limpieza visual
            for (int col = 5; col <= 10; col++) {
                double valor = Double.parseDouble(datos[i][col]);
                String texto = (valor == 0) ? "" : formateadorMoneda.format(valor);
                table.addCell(crearCelda(texto, fontNormal, Element.ALIGN_RIGHT, bg, 3));
            }

            // SALDO (Siempre se muestra y en Negrita)
            table.addCell(crearCelda(formateadorMoneda.format(Double.parseDouble(datos[i][11])), fontNegrita,
                    Element.ALIGN_RIGHT, bg, 3));

            // OBSERVACIONES
            table.addCell(crearCelda(datos[i][12], fontNormal, Element.ALIGN_LEFT, bg, 3));
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
            if (imgMarcaAgua != null) {
                cb.saveState();
                float width = 80;
                float height = 60;
                float x = document.getPageSize().getWidth() - document.rightMargin() - width;
                float y = 15;
                com.lowagie.text.Image imgFooter = com.lowagie.text.Image.getInstance(imgMarcaAgua);
                imgFooter.setAbsolutePosition(x, y);
                imgFooter.scaleToFit(width, height);
                cb.addImage(imgFooter);
                cb.restoreState();
            }
        } catch (Exception e) {
        }
    }
}
