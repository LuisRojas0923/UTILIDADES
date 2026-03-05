import javax.swing.JPanel;
import javax.swing.JTable;

public class MTabla {
    public JTable miTabla = new JTable(); // Agregado para soportar getColumnModel()

    public MTabla(JPanel inPanel, int longitud, int numeroColumnas, int algo, String[] encabezados) {
    }

    public void setAnchoCelda(int idx, int ancho) {
    }

    public void asignaValor(String valor, int fila, int col) {
    } // Agregado

    public void asignarAlturaCelda(int altura) {
    } // Agregado
}
