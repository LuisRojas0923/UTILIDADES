/*
 Proyecto Sistema Integrado de Produccion (SOLID) THC
 
 Clase 				GeneraInforme
 Archivo			GeneraInforme.java
  
 Fecha Creación 	16 de Agosto de 2007
 Autor  			Marlon Ramirez Jaime 1.130.637.509
 email 				marlonrj@gmail.com
*/

import java.io.FileOutputStream;
import java.io.*;
import java.awt.Color;
import javax.swing.*;
import java.awt.*;
import java.awt.image.BufferedImage;

//clases para realizar el informe en pdf
import com.lowagie.text.Document;
import com.lowagie.text.DocumentException;
import com.lowagie.text.Paragraph;
import com.lowagie.text.pdf.PdfWriter;
import com.lowagie.text.pdf.*;
import com.lowagie.text.*;

/**
 * Esta clase genera todos los informes del sistema
 *
 */
public class GeneraInforme {
	
	//atributos
	/** interface con el manejador de base de datos */
	public ConexionDatos miConexionDatos = null;
	
	/** nuevo sistema de impresion centralizado */
	public InformeEstandar miInformeEstandar = null;
	
	/** guarda la fuente por defecto en negrita*/
	public com.lowagie.text.Font fontBold10 = com.lowagie.text.FontFactory.getFont("Helvetica", 10, com.lowagie.text.Font.BOLD,Color.BLACK);
	
	/** guarda la fuente por defecto en negrita*/
	public com.lowagie.text.Font fontBold = com.lowagie.text.FontFactory.getFont("Helvetica", 8, com.lowagie.text.Font.BOLD,Color.BLACK);						
	
	/** guarda la fuente por defecto en negrita*/
	public com.lowagie.text.Font fontBoldSmall = com.lowagie.text.FontFactory.getFont("Helvetica", 8, com.lowagie.text.Font.BOLD,Color.BLACK);						
		
	/** guarda la fuente por defecto en normal*/
	public com.lowagie.text.Font fontNormal = com.lowagie.text.FontFactory.getFont("Helvetica", 8, com.lowagie.text.Font.NORMAL,Color.BLACK);						
	
	/** guarda la fuente por defecto en normal*/
	public com.lowagie.text.Font fontNormalSmall = com.lowagie.text.FontFactory.getFont("Helvetica", 10, com.lowagie.text.Font.NORMAL,Color.BLACK);
	
	/** guarda la fuente por defecto en normal*/
	public com.lowagie.text.Font fontNormalSmall2 = com.lowagie.text.FontFactory.getFont("Helvetica", 4, com.lowagie.text.Font.NORMAL,Color.BLACK);	
	
	/**
	 * constructor por defecto
	 * 
	 * @param inConexionDatos -conexion que se utilizara
	 *
	 */	
	GeneraInforme(ConexionDatos inConexionDatos){
		
		miConexionDatos = inConexionDatos;
		miInformeEstandar = new InformeEstandar(miConexionDatos);
	}
	/**
	 * imprime un plano en un pdf
	 *
	 * @param inImagenPlano -imagen del plano que se imprime
	 *
	 */
	public void exportarPlano(	java.awt.Image inImagenPlano,
								String inTexto){
		Document inDocumento = miInformeEstandar.generarInformeHorizontalBlanco();
		try{
					
			inDocumento.open();
			
			com.lowagie.text.Image imagenPlano = com.lowagie.text.Image.getInstance(inImagenPlano, null);

           	imagenPlano.setAbsolutePosition(40, 60);
			
			imagenPlano.scalePercent(70);

			inDocumento.add(imagenPlano);
			
			java.awt.Image imagen =  (new ImageIcon(FileReaderWriter.folderImages+"logotipo.png")).getImage();
			
			
			InformeEstandar.establecerMarca(imagen, inDocumento, 50, (int)(inDocumento.getPageSize().width()/2- imagen.getWidth(null)*0.5/2.0), 15, inTexto);
			
			inDocumento.close();
		}catch(Exception e){
			e.printStackTrace();
		}
		miInformeEstandar.visualizar();		
	}
	/**
	 * genera un informe estandar utilizado para todos los modulos de consulta
	 *
	 * @param tablaClientes -tabla con los datos
	 *
	 */
	public  PdfPTable generarTablaDatos(	MTabla tablaClientes, 
											String[] encabezadosCliente) throws Exception {
											
		//comienza el algoritmo verdadero
		int nFilas = tablaClientes.miTabla.getSelectedRowCount();
		int nColumnas = tablaClientes.miTabla.getSelectedColumnCount();
		int [] filasSelecionadas = tablaClientes.miTabla.getSelectedRows();
		int [] columnasSelecionadas = tablaClientes.miTabla.getSelectedColumns();
		
		//
		Restriccion.imprimirArreglo(filasSelecionadas);
		Restriccion.imprimirArreglo(columnasSelecionadas);
		//
		
		//calculo de los anchos de las celdas seleccionadas
		int [] anchosCeldas = new int[columnasSelecionadas.length];
		double totalAnchosCeldas = 0.0;
		
		for(int idx1 = 0; idx1 < columnasSelecionadas.length; idx1++){
			anchosCeldas[idx1] = tablaClientes.getAnchoCelda(columnasSelecionadas[idx1]); 
			totalAnchosCeldas += anchosCeldas[idx1];	
		}
		
		System.out.println("Total Anchos de las celdas: "+totalAnchosCeldas);
		
		int [] anchosCeldasPorcent = new int[columnasSelecionadas.length+1];
		anchosCeldasPorcent[0] = 5; 
		for(int idx1 = 0; idx1 < anchosCeldas.length; idx1++){
			anchosCeldasPorcent[idx1+1] = (int)Math.round(((anchosCeldas[idx1] / totalAnchosCeldas) * 100));
		}
		
		System.out.println("Anchos de las celdas: ");
		Restriccion.imprimirArreglo(anchosCeldas);
		
		System.out.println("Porcentaje Anchos de las celdas: ");
		Restriccion.imprimirArreglo(anchosCeldasPorcent);
		
		
		PdfPTable table = new PdfPTable(nColumnas+1);
		table.setWidthPercentage(100);
		table.setWidths(anchosCeldasPorcent);
		table.setHorizontalAlignment(Element.ALIGN_RIGHT);
		
		com.lowagie.text.Font font = com.lowagie.text.FontFactory.getFont("Helvetica", 10, com.lowagie.text.Font.BOLD,Color.BLACK);
		
		PdfPCell cell3 = new PdfPCell(	new Paragraph("",font));
		cell3.setBackgroundColor(new Color(181,181,181));
		cell3.setHorizontalAlignment(Element.ALIGN_CENTER);
		cell3.setVerticalAlignment(Element.ALIGN_MIDDLE);
		table.addCell(cell3);
			
		for(int idx1 = 0; idx1 < nColumnas; idx1++){
			
			PdfPCell cell1 = new PdfPCell(new Paragraph(encabezadosCliente[columnasSelecionadas[idx1]],font));
			cell1.setBackgroundColor(new Color(181,181,181));
			cell1.setHorizontalAlignment(Element.ALIGN_CENTER);
			cell1.setVerticalAlignment(Element.ALIGN_MIDDLE);
			
			table.addCell(cell1);
		}
		
		com.lowagie.text.Font font2 = com.lowagie.text.FontFactory.getFont("Helvetica", 10, com.lowagie.text.Font.NORMAL,Color.BLACK);						
		for(int idx1 = 0; idx1 < nFilas; idx1++){
			PdfPCell cell = new PdfPCell(new Paragraph(""+(idx1+1),font2)	);
			cell.setHorizontalAlignment(Element.ALIGN_CENTER);
			table.addCell(cell);
			for(int idx2 = 0; idx2 < nColumnas; idx2++){
				
				String tmpValor = "";	
				
				try{
					tmpValor = (String)tablaClientes.retornaValor(filasSelecionadas[idx1],columnasSelecionadas[idx2]);
				}catch(Exception e){
					tmpValor = ""+(Boolean)tablaClientes.retornaValor(filasSelecionadas[idx1],columnasSelecionadas[idx2]);
				}
				PdfPCell cell1 = new PdfPCell(new Paragraph(tmpValor,font2)	);
				table.addCell(cell1);
			}
		}
		
		return table;									
	}
	/**
	 * genera un informe estandar utilizado para todos los modulos de consulta
	 *
	 * @param FALTA DEFINIR
	 *
	 */
	public  void generarInformeClientes(	MTabla tablaClientes, 
											String[] encabezadosCliente, 
											String tituloInforme){
		
		
		Document miDocumento = miInformeEstandar.generarInforme();
		try{
					
			miDocumento.open();
			
			miDocumento.add(InformeEstandar.crearTablaLinea("Listado de "+tituloInforme, fontBold10));
			
			miDocumento.add(new Paragraph(	"\n", fontNormal));

			miDocumento.add(new Paragraph(	"\n", fontNormalSmall2));
		
			miDocumento.add(generarTablaDatos(	tablaClientes,
												encabezadosCliente));
			//
			InformeEstandar.establecerMarca(miDocumento);
			//
			
			miDocumento.close();
		}catch(Exception e){
			e.printStackTrace();
		}
		miInformeEstandar.visualizar();
	}
	/**
	 * genera el codigo de barras
	 *
	 * @param inCodigo -codigo a representar
	 *
	 */
	public static java.awt.Image generateBarcode(String inCodigo){
		String codigo = inCodigo;
		double factorEscalacion = 1.5;
		Color colorFondo = new Color(236,233,216);
		int margen = 3;
		colorFondo = Color.white;
		
//		codigo = "m00001";
		
		Barcode128 miBarcode128 = new Barcode128();
		//
//		miBarcode128.setCodeType(com.lowagie.text.pdf.Barcode.CODE128_RAW);
//		miBarcode128.setCodeType(com.lowagie.text.pdf.Barcode.CODE128);
//		miBarcode128.setCodeType(com.lowagie.text.pdf.Barcode.CODE128_UCC);
		//	
		miBarcode128.setCode(codigo);
		
		////////////////////////////////	
		//convirtiendo el codigo a imagen	
		java.awt.Image tmpImagen = miBarcode128.createAwtImage(Color.black, colorFondo);
		
		int ancho = (int)(tmpImagen.getWidth(null) * factorEscalacion);
		int alto = (int)(tmpImagen.getHeight(null) * factorEscalacion);
		
		System.out.println("Alto :"+alto+" ,Ancho :"+ancho);		
		
		//creando la imagen de salida
		BufferedImage imagen = new BufferedImage( ancho + (margen*2), alto + (margen*2) + 13, BufferedImage.TYPE_3BYTE_BGR);	
		
		//editando la imagen de salida
		Graphics2D g = imagen.createGraphics();
		
		//pinta el fondo primero de
		g.setColor(colorFondo);
		g.fillRect(0, 0, imagen.getWidth(null), imagen.getHeight(null));
		
		//dibuja la imagen
		g.drawImage(tmpImagen, 0 + margen, 0 + margen, ancho, alto, null, null);
		
		//dibuja la leyenda del codigo
		g.setColor(java.awt.Color.black);
		g.setFont(new java.awt.Font("Arial", java.awt.Font.PLAIN, 12));
		g.drawString(codigo, margen, alto + margen + 13);
		
		return (java.awt.Image)imagen;
	}		
	/**
	 * genera un informe estandar para los contactos
	 *
	 * @param FALTA DEFINIR
	 *
	 */
	public  void generarInformeContactos(	MTabla tablaContactos, 
											String[] encabezadosContacto, //esta parte es de contactos
											String codigoCliente, 
											String nitCliente,
											String nombreCliente,
											String tituloInforme){

		Document inDocumento = miInformeEstandar.generarInforme();
		try{
					
			inDocumento.open();
			
			inDocumento.add(InformeEstandar.crearTablaLinea("Listado de contactos", fontBold10));
			
			inDocumento.add(new Paragraph(	"\n", fontNormal));
			
			inDocumento.add(new Paragraph("Codigo del "+tituloInforme+":\t"+codigoCliente, fontNormal));
			inDocumento.add(new Paragraph("Nombre del "+tituloInforme+":\t"+nombreCliente, fontNormal));
			inDocumento.add(new Paragraph("Nit del "+tituloInforme+":\t"+nitCliente, fontNormal));

			inDocumento.add(new Paragraph(	"\n", fontNormalSmall2));
		
			inDocumento.add(generarTablaDatos(	tablaContactos,
												encabezadosContacto));
			
			inDocumento.close();
		}catch(Exception e){
			e.printStackTrace();
		}
		miInformeEstandar.visualizar();
	}
	/**
	 * genera un informe estandar de proveedores
	 *
	 * @param FALTA DEFINIR
	 *
	 */
	public  void generarInformeProveedoresMateriaPrima(	MTabla tablaMateriaPrima, 
														String[] encabezados, 
														String codigoMateriaPrima, 
														String descripcionMateriaPrima,
														String tituloInforme){

		Document inDocumento = miInformeEstandar.generarInforme();
		try{
					
			inDocumento.open();
			
			inDocumento.add(InformeEstandar.crearTablaLinea("Listado de proveedores de "+tituloInforme, fontBold10));
			
			inDocumento.add(new Paragraph(	"\n", fontNormal));
			
			inDocumento.add(new Paragraph("Codigo de "+tituloInforme+":\t"+codigoMateriaPrima, fontNormal));
			inDocumento.add(new Paragraph("Descripcion:\t"+descripcionMateriaPrima, fontNormal));

			inDocumento.add(new Paragraph(	"\n", fontNormalSmall2));
		
			inDocumento.add(generarTablaDatos(	tablaMateriaPrima,
												encabezados));
			
			inDocumento.close();
		}catch(Exception e){
			e.printStackTrace();
		}
		miInformeEstandar.visualizar();
	}
	
	/**
	 * genera un informe de programacion de mantenimiento
	 *
	 * @param FALTA DEFINIR
	 *
	 */
	public  void generarInformeProgramacionMTO(	MTabla tablaMateriaPrima, 
												String[] encabezados, 
												String inCodigoMaquina, 
												String tituloInforme){

		Document inDocumento = miInformeEstandar.generarInforme();
		try{
					
			inDocumento.open();
			
			inDocumento.add(InformeEstandar.crearTablaLinea(tituloInforme, fontBold10));
			
			inDocumento.add(new Paragraph(	"\n", fontNormal));
			
			MaquinaInt tmpMaquina = new MaquinaInt(null, miConexionDatos);
			tmpMaquina.cargar("codigo", inCodigoMaquina);
			
			inDocumento.add(new Paragraph("Maquina: "+tmpMaquina.codigo+". "+tmpMaquina.nombre, fontNormal));

			inDocumento.add(new Paragraph(	"\n", fontNormalSmall2));
		
			inDocumento.add(generarTablaDatos(	tablaMateriaPrima,
												encabezados));
			
			inDocumento.close();
		}catch(Exception e){
			e.printStackTrace();
		}
		miInformeEstandar.visualizar();
	}
	
//	/**
//	 * crea una tabla de una fila con fondo gris para titulos
//	 *
//	 * @param inCadena -texto que aparecera en la fila
//	 * @param inFont -fuente que se usara en la tabla
//	 * 
//	 * @return -retorna la tabla creada con la informacion
//	 *
//	 */
//	public PdfPTable crearTablaLinea(String inCadena, com.lowagie.text.Font inFont) throws Exception {
//		//Tabla de una sola fila pra resaltar lineas
//		int nFilas = 1;
//		int nColumnas = 1;
//		float anchoBorde = (float)0.10;
//		PdfPTable tabla = new PdfPTable(nColumnas);
//		tabla.setWidthPercentage(100);
//		PdfPCell celda = new PdfPCell(new Paragraph(inCadena, inFont));
//		celda.setBackgroundColor(new Color(181,181,181));
//		celda.setBorderWidth(anchoBorde);
//		celda.setHorizontalAlignment(Element.ALIGN_CENTER);
//		celda.setVerticalAlignment(Element.ALIGN_MIDDLE);
//		
//		tabla.addCell(celda);
//		
//		return tabla;
//	}
};