import java.awt.Color;
import java.text.DecimalFormat;
import java.text.DecimalFormatSymbols;
import com.lowagie.text.*;
import com.lowagie.text.pdf.*;

/**
 * Clase base para estandarizar la creación de informes en el ERP.
 * Proporciona fuentes, colores y métodos utilitarios para iText 2.1.7.
 */
public abstract class BaseInforme extends PdfPageEventHelper {

    // Colores Corporativos
    public static final Color AZUL_NAVY = new Color(0, 32, 96);
    public static final Color GRIS_CLARO = new Color(242, 242, 242);
    public static final Color GRIS_BORDE = new Color(166, 166, 166);

    // Fuentes Estándar
    public Font fontTitulo = FontFactory.getFont("Helvetica", 12, Font.BOLD, AZUL_NAVY);
    public Font fontSubtitulo = FontFactory.getFont("Helvetica", 10, Font.BOLD, AZUL_NAVY);
    public Font fontNegritaBlancas = FontFactory.getFont("Helvetica", 8, Font.BOLD, Color.WHITE);
    public Font fontNormal = FontFactory.getFont("Helvetica", 8, Font.NORMAL, Color.BLACK);
    public Font fontNegrita = FontFactory.getFont("Helvetica", 8, Font.BOLD, Color.BLACK);
    public Font fontMini = FontFactory.getFont("Helvetica", 7, Font.NORMAL, Color.BLACK);

    // Formateadores
    protected DecimalFormat formateadorMoneda;

    public BaseInforme() {
        DecimalFormatSymbols simbolos = new DecimalFormatSymbols();
        simbolos.setDecimalSeparator(',');
        simbolos.setGroupingSeparator('.');
        formateadorMoneda = new DecimalFormat("$ #,###", simbolos);
    }

    /**
     * Crea una celda con estilo predefinido.
     */
    protected PdfPCell crearCelda(String texto, Font fuente, int alineacion, Color fondo, float padding) {
        PdfPCell celda = new PdfPCell(new Phrase(texto, fuente));
        celda.setHorizontalAlignment(alineacion);
        celda.setVerticalAlignment(Element.ALIGN_MIDDLE);
        if (fondo != null)
            celda.setBackgroundColor(fondo);
        celda.setPadding(padding);
        celda.setBorderColor(GRIS_BORDE);
        return celda;
    }

    protected PdfPCell crearCeldaSinBorde(String texto, Font fuente, int alineacion) {
        PdfPCell celda = new PdfPCell(new Phrase(texto, fuente));
        celda.setHorizontalAlignment(alineacion);
        celda.setBorder(PdfPCell.NO_BORDER);
        return celda;
    }

    protected PdfPCell crearCelda(String texto, Font fuente, int alineacion, Color fondo, float padding, int colspan) {
        PdfPCell celda = crearCelda(texto, fuente, alineacion, fondo, padding);
        celda.setColspan(colspan);
        return celda;
    }

    /**
     * Método abstracto que cada informe debe implementar para su contenido
     * principal.
     */
    public abstract void generar(String parametroId) throws Exception;
}
