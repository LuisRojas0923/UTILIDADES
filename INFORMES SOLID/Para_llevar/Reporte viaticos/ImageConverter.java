import java.awt.Image;
import java.io.ByteArrayInputStream;
import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;

public class ImageConverter {
    public static Image convertToImage(Object representacion) {
        if (representacion == null)
            return null;

        try {
            byte[] bytes = null;

            if (representacion instanceof byte[]) {
                bytes = (byte[]) representacion;
            } else if (representacion instanceof String) {
                String s = ((String) representacion).trim();
                // Limpiar comillas si existen
                if (s.startsWith("\"") && s.endsWith("\"")) {
                    s = s.substring(1, s.length() - 1);
                }

                String[] parts = s.split(",");
                bytes = new byte[parts.length];
                for (int i = 0; i < parts.length; i++) {
                    try {
                        // Intentar parseo directo como byte (soporta -128 a 127)
                        bytes[i] = Byte.parseByte(parts[i].trim());
                    } catch (NumberFormatException nfe) {
                        try {
                            // Si falla (ej: "129"), intentar como short y cast (0-255 sin signo)
                            short val = Short.parseShort(parts[i].trim());
                            bytes[i] = (byte) val;
                        } catch (Exception e2) {
                            System.out.println("Error parsing byte: " + parts[i]);
                        }
                    }
                }
            }

            if (bytes != null && bytes.length > 0) {
                return ImageIO.read(new ByteArrayInputStream(bytes));
            }
        } catch (Exception e) {
            System.out.println("Error converting image: " + e.getMessage());
            e.printStackTrace();
        }
        return null;
    }
}
