/*
 Proyecto Solid refridcol
 
 Clase 				GeneraCertificadoLaboral
 Archivo			GeneraCertificadoLaboral.java
  
 Fecha Creación 	10 de Febrero de 2025
 Autor  			Marlon Ramirez Jaime 1.130.637.509
 email 				marlonrj@gmail.com
*/

import java.io.FileOutputStream;
import java.io.*;
import java.awt.Color;
import javax.swing.*;
import java.awt.*;
import java.util.*;
import java.awt.image.BufferedImage;

//clases para realizar el informe en pdf
import com.lowagie.text.Document;
import com.lowagie.text.DocumentException;
import com.lowagie.text.Paragraph;
import com.lowagie.text.pdf.PdfWriter;
import com.lowagie.text.pdf.*;
import com.lowagie.text.*;
import java.sql.*;
import java.text.DecimalFormat;
import java.text.DecimalFormatSymbols;

/**
 * Esta clase genera el formato de cotizacion de venta
 *
 */
public class GeneraCertificadoLaboral extends PdfPageEventHelper {
	
	//atributos
	/** interface con el manejador de base de datos */
	public ConexionDatos miConexionDatos = null;
	
	/** manejador de archivos */
	public FileReaderWriter miFileReaderWriter = null;
	
	/** guarda la direccion donde se debe almacenar el informe generado*/
	public String dirInforme = "";
	
	/** guarda la fuente por defecto en negrita*/
	public com.lowagie.text.Font fontBold = com.lowagie.text.FontFactory.getFont("Arial", 11, com.lowagie.text.Font.BOLD,Color.BLACK);	
		
	/** guarda la fuente por defecto en negrita*/
	public com.lowagie.text.Font fontBoldTitulo = com.lowagie.text.FontFactory.getFont("Arial", 18, com.lowagie.text.Font.BOLD,new java.awt.Color(166,166,166));						
	
	/** guarda la fuente por defecto en negrita*/
	public com.lowagie.text.Font fontBoldSmall = com.lowagie.text.FontFactory.getFont("Arial", 8, com.lowagie.text.Font.BOLD,Color.BLACK);
	
	/** guarda la fuente por defecto en negrita*/
	public com.lowagie.text.Font fontSaltoNotas = com.lowagie.text.FontFactory.getFont("Arial", 6, com.lowagie.text.Font.NORMAL,Color.BLACK);
	
	/** guarda la fuente por defecto en negrita*/
	public com.lowagie.text.Font fontSaltoOT = com.lowagie.text.FontFactory.getFont("Arial", 3, com.lowagie.text.Font.NORMAL,Color.BLACK);							
		
	/** guarda la fuente por defecto en normal*/
	public com.lowagie.text.Font fontNormal = com.lowagie.text.FontFactory.getFont("Arial", 11, com.lowagie.text.Font.NORMAL,Color.BLACK);						
	
	/** guarda la fuente por defecto en normal*/
	public com.lowagie.text.Font fontNormalSmall = com.lowagie.text.FontFactory.getFont("Arial", 8, com.lowagie.text.Font.NORMAL,Color.BLACK);
	
		/** guarda la fuente por defecto en normal*/
	public com.lowagie.text.Font fontNormalSmallGris = com.lowagie.text.FontFactory.getFont("Helvetica", 8, com.lowagie.text.Font.NORMAL,Color.GRAY);								
	
	/** guarda el indice de la fila actual que se esta imprimiendo*/
	int gFilaActual = 0;
	
	/** permite cargar y almacenar la informacion de la empresa*/
	public Empresa tmpEmpresa = null;
	
	/** permite cargar y almacenar la informacion de la cotizacion*/
	Contrato tmpContrato = null;

	Establecimiento tmpPersonal = null;

	
	/** guarda la imagen que se usa como membrete del documento*/
	public java.awt.Image imagenMembrete = null;
	
	/** guarda la imagen que se usa como membrete del documento*/
	public java.awt.Image imagenFirma = null;
	
	/** guarda la imagen que se usa como membrete del documento*/
	public java.awt.Image imagenPortada = null;
	
	/** formateador campos de valor */
	DecimalFormatSymbols simbolos = new DecimalFormatSymbols();
	DecimalFormatSymbols simbolos2 = new DecimalFormatSymbols();
	DecimalFormat formateador = null;
	DecimalFormat formateadorPesos = null;
			
	/**
	 * constructor por defecto
	 *
	 */	
	GeneraCertificadoLaboral(ConexionDatos inConexionDatos){
		miConexionDatos = inConexionDatos;
		
		simbolos.setDecimalSeparator('.');
		simbolos.setGroupingSeparator(',');
		simbolos2.setDecimalSeparator(',');
		simbolos2.setGroupingSeparator('.');
		formateador = new DecimalFormat("#,###", simbolos2);
		formateadorPesos = new DecimalFormat("$ #,###.00", simbolos);
		
		miFileReaderWriter = new FileReaderWriter(miConexionDatos);
		miFileReaderWriter.cargarParametrosInforme();
		miFileReaderWriter.cargarParametrosCotizacionVenta();
		miFileReaderWriter.cargarPiePaginaInforme();
		
		tmpEmpresa = new Empresa(null, miConexionDatos);
		tmpEmpresa.cargar("codigo", "1");
		
		Imagen tmpImagen = new Imagen(
										null,
										miConexionDatos
										);
        
    	tmpImagen.cargar("codigo", "3");
    	
    	try{
    		imagenMembrete = ImageConverter.convertToImage(tmpImagen.representacion);
    	}catch(Exception e){
    		BufferedImage imagen = new BufferedImage( 200, 200, BufferedImage.TYPE_INT_ARGB);
    		Graphics2D g2d = imagen.createGraphics();
    		g2d.fillRect(0, 0, imagen.getWidth(null), imagen.getHeight(null));
    		imagenMembrete = (java.awt.Image) imagen;
    	}
    	
    	
    	Imagen tmpImagen2 = new Imagen(
										null,
										miConexionDatos
										);
        
    	tmpImagen2.cargar("codigo", "1");
    	
    	try{
    		imagenFirma = ImageConverter.convertToImage(tmpImagen2.representacion);
    	}catch(Exception e){
    		BufferedImage imagen = new BufferedImage( 200, 200, BufferedImage.TYPE_INT_ARGB);
    		Graphics2D g2d = imagen.createGraphics();
    		g2d.fillRect(0, 0, imagen.getWidth(null), imagen.getHeight(null));
    		imagenFirma = (java.awt.Image) imagen;
    	}
	}
	////////////////////////////////////////////////////////////////////////
	///////////////////GENERACION DE LA COTIZACION DE VENTA//////////////////
	/**
	 * genera el encabezado de la ficha del producto (Datos generales)
	 *
	 * @param inVectorDatosCotizacion -vector con los datos de la cotizacion a imprimir
	 * @param inTablaDetallesCotizacion -MTabla con los datos de los detalles de cotizacion a imprimir
	 *
	 */
	public  void generarCertificado(	String InCodigoContrato){ //true consolidada
		try{
		
//			if(SOLID.versionWeb){
//
//				JFileChooser selector = new JFileChooser();		
//				File archivo1 = new File("Certificado laboral "+InCodigoContrato+" - "+(int)(Math.random()*1000)+".pdf");
//				selector.setSelectedFile(archivo1);
//				selector.showSaveDialog(null);
//				
//				File archivo2 = selector.getSelectedFile();
//				String nombreArchivo = archivo2.getPath();
//				dirInforme = archivo2.getPath();
//				
//			}else{
				dirInforme = FileReaderWriter.establecerNombreArchivo(miConexionDatos, "Certificado laboral "+InCodigoContrato+" - "+(Math.random()*1000));
//			}
			
			
			//creacion del documento y establecimiento de los margenes
			Document miDocumento = new Document(PageSize.LETTER,50,60,70,50);
			
			//creacion de la fecha
			com.lowagie.text.Font fuenteFecha = com.lowagie.text.FontFactory.getFont("Arial", 11, com.lowagie.text.Font.NORMAL,Color.BLACK);

			//creacion del titulo
			com.lowagie.text.Font fuenteTitulos = com.lowagie.text.FontFactory.getFont("Arial", 11, com.lowagie.text.Font.BOLD,Color.BLACK);
			
			//armado del documento
			PdfWriter writer = PdfWriter.getInstance(miDocumento, new FileOutputStream(dirInforme));
			writer.setPageEvent(this);
			
			
			
			/////////carga de los datos de la orden de trabajo
			tmpContrato = new Contrato(	null,
														miConexionDatos);
			tmpContrato.cargar("Codigo", InCodigoContrato);
			
			
			
			
			//personal
			tmpPersonal = new Establecimiento(	null,
													miConexionDatos);
			tmpPersonal.cargar("Codigo", tmpContrato.establecimiento);
			/////////
			
			
			
			//aqui se determina si se imprime con el codigo o con referencia
			Paragraph numeroCotizacion = null;
			
			


			
			miDocumento.open();
			
			int anchoTablaItems = (int)(miDocumento.getPageSize().width() - miDocumento.leftMargin() - miDocumento.rightMargin());
			
			int numeroPasada = 0;
			while(numeroPasada==0){
//			while((gFilaActual != inTablaDetallesCotizacion.miModeloTabla.datos.length) || numeroPasada==0){	
				numeroPasada++; 
				
				

				
				
				System.out.println("indice actual: "+gFilaActual);
			
			
				//ok Orden de trabajo
					
//				miDocumento.add(InformeEstandar.crearTablaLineaIncoCentrada("PROVEEDOR:", fontBold));
//				miDocumento.add(new Paragraph(	"\n", fontSaltoOT));	
					
		
				
				Calendar fecha = Calendar.getInstance();

				int fechaDiaA = fecha.get(Calendar.DAY_OF_MONTH);
				int fechaMesA = fecha.get(Calendar.MONTH);
				int fechaYearA = fecha.get(Calendar.YEAR);
				
				Phrase tmpTextoNormal = new Phrase(	"\n\n"+tmpEmpresa.ciudad+", "+fechaDiaA+" de "+SOLID.convertirMes(fechaMesA)+" de "+fechaYearA, fontNormal);
				Paragraph tmpParrafo = new Paragraph();
				
				tmpParrafo.add(tmpTextoNormal);
//				tmpParrafo.setAlignment(Element.ALIGN_JUSTIFIED);
				miDocumento.add(tmpParrafo);
				
				miDocumento.add(new Paragraph(	"\n\n\n\n", fontNormal));
				
				
				tmpTextoNormal = new Phrase(	"Señores\n", fontNormal);
				Phrase tmpTextoNegrilla = new Phrase(	"A quien interese\n\n\n", fontBold);
				Phrase tmpTextoNormal2 = new Phrase(	"Asunto: Certificado laboral "+tmpPersonal.nombre+"\n\n", fontBold);
				tmpParrafo = new Paragraph();
				
				tmpParrafo.add(tmpTextoNormal);
				tmpParrafo.add(tmpTextoNegrilla);
				tmpParrafo.add(tmpTextoNormal2);
//				tmpParrafo.setAlignment(Element.ALIGN_JUSTIFIED);
				miDocumento.add(tmpParrafo);
				
				miDocumento.add(new Paragraph(	"\n", fontNormal));
				
				StringTokenizer miTK = new StringTokenizer(tmpContrato.fechaInicio,"-");
		    	int fechaYear = Integer.parseInt(miTK.nextToken());
				int fechaMes = Integer.parseInt(miTK.nextToken());
				int fechaDia = Integer.parseInt(miTK.nextToken());
				
				NumberToText miConvertidor = new NumberToText(fechaDia);
				tmpTextoNormal = new Phrase(	"Por medio de la presente ", fontNormal);
				Phrase tmpTextoNegritaInco = new Phrase(	tmpEmpresa.nombreEmpresa, fontBold);
				tmpTextoNormal2 = new Phrase(	" hace constar que el(la) señor(a) ", fontNormal);
				
				tmpTextoNegrilla = new Phrase(	tmpPersonal.nombre, fontBold);
				Phrase tmpTextoNormal3 = new Phrase(	" identificado(a) con cédula de ciudadanía No. "+formateador.format(Double.parseDouble(tmpPersonal.nroCedula))+""
												+" labora en nuestra compañia desde el "+fechaDia+" de "+SOLID.convertirMes(fechaMes).toUpperCase()+" del año "+fechaYear
												+" con un tipo de contrato "	  , fontNormal);
				Phrase tmpTextoNegritaTipo = new Phrase(	tmpContrato.tipo	  , fontBold);
				
				Phrase tmpTextoNormal4 = new Phrase(	", desempeñando el cargo de "	  , fontNormal);
				
				Phrase tmpTextoNegritaCargo = new Phrase(	tmpPersonal.cargo, fontBold);																
				
				/////CONSULTAR MRJ2025 EL MAS RECIENTE SALARIO
				/////
				Phrase tmpTextoNormal5 = new Phrase(", devengando un salario basico mensual de "+formateadorPesos.format(Double.parseDouble("0.0"))
												+" ("+miConvertidor.convertirLetras(Integer.parseInt("0")).trim().toUpperCase()+" MCTE)."	  , fontNormal);
				
				tmpParrafo = new Paragraph();
				
				tmpParrafo.add(tmpTextoNormal);
				tmpParrafo.add(tmpTextoNegritaInco);
				tmpParrafo.add(tmpTextoNormal2);
				tmpParrafo.add(tmpTextoNegrilla);
				tmpParrafo.add(tmpTextoNormal3);
				tmpParrafo.add(tmpTextoNegritaTipo);
				tmpParrafo.add(tmpTextoNormal4);
				tmpParrafo.add(tmpTextoNegritaCargo);
				tmpParrafo.add(tmpTextoNormal5);
				tmpParrafo.setAlignment(Element.ALIGN_JUSTIFIED);
				miDocumento.add(tmpParrafo);
				
				
				
				tmpTextoNormal = new Phrase(	"\n\nLa presente se expide a solicitud del interesado el día "+fechaDiaA+" ("+miConvertidor.convertirLetras(fechaDiaA).trim().toUpperCase()+") del mes de "+SOLID.convertirMes(fechaMes).toUpperCase()+" del año "+fechaYear+".", fontNormal);
				
				tmpTextoNormal2 = new Phrase(	"\n\n\nCordialmente,", fontNormal);
				tmpTextoNegrilla = new Phrase(	"\n\n\n\n\n_____________________________\nStephanie Dumancely\n", fontBold);
				tmpTextoNormal3 = new Phrase(	"Jefe Gestión Humana\nTeléfono +57 311 385 9927\njefegestionhumana@refridcol.com", fontNormal);
				
				tmpParrafo = new Paragraph();
				
				tmpParrafo.add(tmpTextoNormal);
				tmpParrafo.add(tmpTextoNormal2);
				tmpParrafo.add(tmpTextoNegrilla);
				tmpParrafo.add(tmpTextoNormal3);
				miDocumento.add(tmpParrafo);
				
				////
				miDocumento.add(new Paragraph(	"\n", fontNormal));	
				miDocumento.add(new Paragraph(	"\n", fontNormal));	
				
//				miDocumento.add(crearTablaFirmas( anchoTablaItems));

	
				miDocumento.newPage();
			}			
			miDocumento.close();
		}catch(Exception e){
			e.printStackTrace();
		}
		
//////////		//abre el archivo de salida
//////////		File tmpFile = new File(dirInforme);
//////////		if(tmpFile.exists()){
//////////			try{
//////////				Runtime.getRuntime().exec(miFileReaderWriter.dirVisualizador+" \""+dirInforme+" \"" );
//////////			}catch(Exception e){
//////////				e.printStackTrace();
//////////				JOptionPane.showMessageDialog(null,"No se puede abrir el archivo, verifique la ruta de la aplicación para leer archivos PDF");
//////////			}
//////////		}
		
//		if(FORCE.versionWeb){
//			//no se viasualiza porque se define por webswing
//		}else{
			//abre el archivo de salida
			File tmpFile = new File(dirInforme);
			if(tmpFile.exists()){
				try{
					
					
					Runtime.getRuntime().exec(" cmd /c \""+dirInforme+" \"" );
				}catch(Exception e){
					e.printStackTrace();
					JOptionPane.showMessageDialog(null,"No se puede abrir el archivo, verifique la ruta de la aplicación para leer archivos PDF");
				}
			}
//		}
	}
	/**
	 * crea la tabla items sencilla
	 *
	 */
	public PdfPTable crearTablaFirmas(	
														int inAncho
														) throws Exception {

		//tabla con los datos para la fila 2 de la tablaEx
		int nFilasDatos = 13;
		int nColumnasDatos = 2;
		float anchoBorde = (float)0.10;
		float anchoBordeInterno = (float)0.10;
		
		PdfPTable tablaDatos = new PdfPTable(nColumnasDatos);
	
		
		int[] anchosTabla = {50,50};
		tablaDatos.setWidths(anchosTabla);
		tablaDatos.setWidthPercentage(100);
		
		//estableciendo ancho de la tabla
		tablaDatos.setTotalWidth(inAncho);
		
		/////titulo cliente
		PdfPCell cell1 = new PdfPCell(new Paragraph("EL EMPLEADOR",fontBold));
		cell1.setBorderWidth((float)0.0);
//		cell1.setBackgroundColor(new Color(217,217,217));
//		cell1.setBorderColor(new java.awt.Color(166,166,166));
		tablaDatos.addCell(cell1);
		
		/////titulo cliente
		cell1 = new PdfPCell(new Paragraph("EL TRABAJADOR",fontBold));
		cell1.setBorderWidth((float)0.0);
//		cell1.setBackgroundColor(new Color(217,217,217));
//		cell1.setBorderColor(new java.awt.Color(166,166,166));
		tablaDatos.addCell(cell1);
			
		/////titulo ubicacion
	    cell1 = new PdfPCell(new Paragraph("\n\n\n\n\n\n_______________________________",fontBold));
		cell1.setBorderWidth((float)0.0);
//		cell1.setBackgroundColor(new Color(217,217,217));
//		cell1.setBorderColor(new java.awt.Color(166,166,166));
		tablaDatos.addCell(cell1);
		
	    cell1 = new PdfPCell(new Paragraph("\n\n\n\n\n\n_______________________________",fontBold));
		cell1.setBorderWidth((float)0.0);
//		cell1.setBackgroundColor(new Color(217,217,217));
//		cell1.setBorderColor(new java.awt.Color(166,166,166));
		tablaDatos.addCell(cell1);	
			
		/////titulo contacto
		cell1 = new PdfPCell(new Paragraph("Camilo Andrés Rodríguez Beltrán",fontBold));
		cell1.setBorderWidth((float)0.0);
//		cell1.setBackgroundColor(new Color(217,217,217));
//		cell1.setBorderColor(new java.awt.Color(166,166,166));
		tablaDatos.addCell(cell1);
		
		//dato CONTACTO
		cell1 = new PdfPCell(new Paragraph(tmpPersonal.nombre,fontBold));
		cell1.setBorderWidth((float)0.0);
//		cell1.setBorderColor(new java.awt.Color(166,166,166));
		tablaDatos.addCell(cell1);	
			
		/////titulo RECURSO HUMANO
		cell1 = new PdfPCell(new Paragraph("C.C. 1.019.030.197",fontNormal));
		cell1.setBorderWidth((float)0.0);
//		cell1.setBackgroundColor(new Color(217,217,217));
//		cell1.setBorderColor(new java.awt.Color(166,166,166));
		tablaDatos.addCell(cell1);
		
		//dato RECURSO HUMANO
	
		
		cell1 = new PdfPCell(new Paragraph("C.C. "+formateador.format(Double.parseDouble(tmpPersonal.nroCedula)),fontNormal));
		cell1.setBorderWidth((float)0.0);
//		cell1.setBorderColor(new java.awt.Color(166,166,166));
		tablaDatos.addCell(cell1);	
			
		/////titulo NUMERO COTIZACION
		cell1 = new PdfPCell(new Paragraph("Representante Legal",fontNormal));
		cell1.setBorderWidth((float)0.0);
//		cell1.setBackgroundColor(new Color(217,217,217));
//		cell1.setBorderColor(new java.awt.Color(166,166,166));
		tablaDatos.addCell(cell1);
		
		//dato NUMERO COTIZACION
		cell1 = new PdfPCell(new Paragraph("",fontNormal));
		cell1.setBorderWidth((float)0.0);
//		cell1.setBorderColor(new java.awt.Color(166,166,166));
		tablaDatos.addCell(cell1);	
			
		/////titulo NUMERO COTIZACION
		cell1 = new PdfPCell(new Paragraph(tmpEmpresa.nombreEmpresa,fontNormal));
		cell1.setBorderWidth((float)0.0);
//		cell1.setBackgroundColor(new Color(217,217,217));
//		cell1.setBorderColor(new java.awt.Color(166,166,166));
		tablaDatos.addCell(cell1);
		
		//dato NUMERO COTIZACION
		cell1 = new PdfPCell(new Paragraph("",fontNormal));
		cell1.setBorderWidth((float)0.0);
//		cell1.setBorderColor(new java.awt.Color(166,166,166));
		tablaDatos.addCell(cell1);	
		
		return tablaDatos;
	}
	/**
	 * crea la tabla items sencilla
	 *
	 */
	public PdfPTable crearTablaInformacionEmpleado(	
														int inAncho
														) throws Exception {

		//tabla con los datos para la fila 2 de la tablaEx
		int nFilasDatos = 13;
		int nColumnasDatos = 2;
		float anchoBorde = (float)0.10;
		float anchoBordeInterno = (float)0.10;
		
		PdfPTable tablaDatos = new PdfPTable(nColumnasDatos);
	
		
		int[] anchosTabla = {28,72};
		tablaDatos.setWidths(anchosTabla);
		tablaDatos.setWidthPercentage(100);
		
		//estableciendo ancho de la tabla
		tablaDatos.setTotalWidth(inAncho);
		
		/////titulo cliente
		PdfPCell cell1 = new PdfPCell(new Paragraph("Nombre del empleador:",fontBold));
		cell1.setBorderWidth((float)0.0);
//		cell1.setBackgroundColor(new Color(217,217,217));
//		cell1.setBorderColor(new java.awt.Color(166,166,166));
		tablaDatos.addCell(cell1);
		
		//dato cliente
		cell1 = new PdfPCell(new Paragraph(tmpEmpresa.nombreEmpresa,fontNormal));
		cell1.setBorderWidth((float)0.0);
//		cell1.setBorderColor(new java.awt.Color(166,166,166));
		tablaDatos.addCell(cell1);	
			
		/////titulo ubicacion
	    cell1 = new PdfPCell(new Paragraph("Identificación:",fontBold));
		cell1.setBorderWidth((float)0.0);
//		cell1.setBackgroundColor(new Color(217,217,217));
//		cell1.setBorderColor(new java.awt.Color(166,166,166));
		tablaDatos.addCell(cell1);
		
		//dato UBICACION
		cell1 = new PdfPCell(new Paragraph(tmpEmpresa.nit,fontNormal));
		cell1.setBorderWidth((float)0.0);
//		cell1.setBorderColor(new java.awt.Color(166,166,166));
		tablaDatos.addCell(cell1);	
			
		/////titulo contacto
		cell1 = new PdfPCell(new Paragraph("Dirección:",fontBold));
		cell1.setBorderWidth((float)0.0);
//		cell1.setBackgroundColor(new Color(217,217,217));
//		cell1.setBorderColor(new java.awt.Color(166,166,166));
		tablaDatos.addCell(cell1);
		
		//dato CONTACTO
		cell1 = new PdfPCell(new Paragraph(tmpEmpresa.direccion,fontNormal));
		cell1.setBorderWidth((float)0.0);
//		cell1.setBorderColor(new java.awt.Color(166,166,166));
		tablaDatos.addCell(cell1);	
			
		/////titulo RECURSO HUMANO
		cell1 = new PdfPCell(new Paragraph("Nombre del trabajador:",fontBold));
		cell1.setBorderWidth((float)0.0);
//		cell1.setBackgroundColor(new Color(217,217,217));
//		cell1.setBorderColor(new java.awt.Color(166,166,166));
		tablaDatos.addCell(cell1);
		
		//dato RECURSO HUMANO
	
		
		cell1 = new PdfPCell(new Paragraph(tmpPersonal.nombre,fontNormal));
		cell1.setBorderWidth((float)0.0);
//		cell1.setBorderColor(new java.awt.Color(166,166,166));
		tablaDatos.addCell(cell1);	
			
		/////titulo NUMERO COTIZACION
		cell1 = new PdfPCell(new Paragraph("Identificación:",fontBold));
		cell1.setBorderWidth((float)0.0);
//		cell1.setBackgroundColor(new Color(217,217,217));
//		cell1.setBorderColor(new java.awt.Color(166,166,166));
		tablaDatos.addCell(cell1);
		
		//dato NUMERO COTIZACION
		cell1 = new PdfPCell(new Paragraph(tmpPersonal.nroCedula,fontNormal));
		cell1.setBorderWidth((float)0.0);
//		cell1.setBorderColor(new java.awt.Color(166,166,166));
		tablaDatos.addCell(cell1);	
			
		/////titulo LINEA DE SERVICIO
		cell1 = new PdfPCell(new Paragraph("Fecha de nacimiento:",fontBold));
		cell1.setBorderWidth((float)0.0);
//		cell1.setBackgroundColor(new Color(217,217,217));
//		cell1.setBorderColor(new java.awt.Color(166,166,166));
		tablaDatos.addCell(cell1);
		
		//dato LINEA DE SERVICIO
		cell1 = new PdfPCell(new Paragraph(tmpPersonal.fechaNacimiento,fontNormal));
		cell1.setBorderWidth((float)0.0);
//		cell1.setBorderColor(new java.awt.Color(166,166,166));
		tablaDatos.addCell(cell1);	
			
		
			/////titulo LINEA DE SERVICIO
		cell1 = new PdfPCell(new Paragraph("Dirección:",fontBold));
		cell1.setBorderWidth((float)0.0);
//		cell1.setBackgroundColor(new Color(217,217,217));
//		cell1.setBorderColor(new java.awt.Color(166,166,166));
		tablaDatos.addCell(cell1);
		
		//dato LINEA DE SERVICIO
		cell1 = new PdfPCell(new Paragraph(tmpPersonal.direccionResidencia,fontNormal));
		cell1.setBorderWidth((float)0.0);
//		cell1.setBorderColor(new java.awt.Color(166,166,166));
		tablaDatos.addCell(cell1);	
		
		
				/////titulo LINEA DE SERVICIO
		cell1 = new PdfPCell(new Paragraph("Ciudad:",fontBold));
		cell1.setBorderWidth((float)0.0);
//		cell1.setBackgroundColor(new Color(217,217,217));
//		cell1.setBorderColor(new java.awt.Color(166,166,166));
		tablaDatos.addCell(cell1);
		
		//dato LINEA DE SERVICIO
		cell1 = new PdfPCell(new Paragraph(tmpPersonal.ciudadResidencia,fontNormal));
		cell1.setBorderWidth((float)0.0);
//		cell1.setBorderColor(new java.awt.Color(166,166,166));
		tablaDatos.addCell(cell1);	
			
					/////titulo LINEA DE SERVICIO
		cell1 = new PdfPCell(new Paragraph("Correo electrónico:",fontBold));
		cell1.setBorderWidth((float)0.0);
//		cell1.setBackgroundColor(new Color(217,217,217));
//		cell1.setBorderColor(new java.awt.Color(166,166,166));
		tablaDatos.addCell(cell1);
		
		//dato LINEA DE SERVICIO
		cell1 = new PdfPCell(new Paragraph(tmpPersonal.correoPersonal,fontNormal));
		cell1.setBorderWidth((float)0.0);
//		cell1.setBorderColor(new java.awt.Color(166,166,166));
		tablaDatos.addCell(cell1);	
		
							/////titulo LINEA DE SERVICIO
		cell1 = new PdfPCell(new Paragraph("Celular:",fontBold));
		cell1.setBorderWidth((float)0.0);
//		cell1.setBackgroundColor(new Color(217,217,217));
//		cell1.setBorderColor(new java.awt.Color(166,166,166));
		tablaDatos.addCell(cell1);
		
		//dato LINEA DE SERVICIO
		cell1 = new PdfPCell(new Paragraph(tmpPersonal.telefono,fontNormal));
		cell1.setBorderWidth((float)0.0);
//		cell1.setBorderColor(new java.awt.Color(166,166,166));
		tablaDatos.addCell(cell1);	
			
		/////titulo LINEA DE SERVICIO
		cell1 = new PdfPCell(new Paragraph("Cargo:",fontBold));
		cell1.setBorderWidth((float)0.0);
//		cell1.setBackgroundColor(new Color(217,217,217));
//		cell1.setBorderColor(new java.awt.Color(166,166,166));
		tablaDatos.addCell(cell1);
		
		//dato LINEA DE SERVICIO
		cell1 = new PdfPCell(new Paragraph(tmpPersonal.cargo,fontNormal));
		cell1.setBorderWidth((float)0.0);
//		cell1.setBorderColor(new java.awt.Color(166,166,166));
		tablaDatos.addCell(cell1);		
		
				/////titulo LINEA DE SERVICIO
		cell1 = new PdfPCell(new Paragraph("Salario:",fontBold));
		cell1.setBorderWidth((float)0.0);
//		cell1.setBackgroundColor(new Color(217,217,217));
//		cell1.setBorderColor(new java.awt.Color(166,166,166));
		tablaDatos.addCell(cell1);
		
		//dato LINEA DE SERVICIO
		cell1 = new PdfPCell(new Paragraph("0.0",fontNormal));
		cell1.setBorderWidth((float)0.0);
//		cell1.setBorderColor(new java.awt.Color(166,166,166));
		tablaDatos.addCell(cell1);	
			
				/////titulo LINEA DE SERVICIO
		cell1 = new PdfPCell(new Paragraph("Periodicidad de Pago:",fontBold));
		cell1.setBorderWidth((float)0.0);
//		cell1.setBackgroundColor(new Color(217,217,217));
//		cell1.setBorderColor(new java.awt.Color(166,166,166));
		tablaDatos.addCell(cell1);
		
		//dato LINEA DE SERVICIO
		cell1 = new PdfPCell(new Paragraph("Quincenal",fontNormal));
		cell1.setBorderWidth((float)0.0);
//		cell1.setBorderColor(new java.awt.Color(166,166,166));
		tablaDatos.addCell(cell1);	
			
					/////titulo LINEA DE SERVICIO
		cell1 = new PdfPCell(new Paragraph("Fecha de Inicio:",fontBold));
		cell1.setBorderWidth((float)0.0);
//		cell1.setBackgroundColor(new Color(217,217,217));
//		cell1.setBorderColor(new java.awt.Color(166,166,166));
		tablaDatos.addCell(cell1);
		
		//dato LINEA DE SERVICIO
		cell1 = new PdfPCell(new Paragraph(tmpContrato.fechaInicio,fontNormal));
		cell1.setBorderWidth((float)0.0);
//		cell1.setBorderColor(new java.awt.Color(166,166,166));
		tablaDatos.addCell(cell1);			
		
		return tablaDatos;
	}
	



		
	/**
	 * crea la tabla total de orden de compra
	 *
	 */
	public PdfPTable crearTablaTOTALOrdenCompra(	String inCodigoOrdenCompra,
																int inAncho ) throws Exception {

		ResultSet resultado;		
		resultado = miConexionDatos.funcionConsultar(	"select sumar(sum(lineaOrdenCompra.cantidad * lineaOrdenCompra.valorUnitario),0) "
														+" from lineaOrdenCompra "
														+" where "
														+" lineaOrdenCompra.ordenCompra = '"+inCodigoOrdenCompra+"' ");
																			
		//convierte el resultado en un arreglo
		String datosTotal[][] = ConexionDatos.armarArreglo(resultado);
		
		Double valorTotal = 0.0;
		valorTotal = Double.parseDouble(datosTotal[0][0]);
		
																					
		//tabla con los datos para la fila 2 de la tablaEx
		int nFilasDatos = 1;
		int nColumnasDatos = 3;
		float anchoBorde = (float)0.10;
		float anchoBordeInterno = (float)0.10;
		
		PdfPTable tablaDatos = new PdfPTable(nColumnasDatos);
		
//////		System.out.println("filas:"+nFilasDatos+", Columnas:"+nColumnasDatos);
//////		
//////		String[] nombresEncabezado = {	"Item" 	};	
//////		
//////		int[] anchosTabla = {45,30,25};
//////		tablaDatos.setWidths(anchosTabla);
//////		tablaDatos.setWidthPercentage(100);
//////		
//////		
//////		//estableciendo ancho de la tabla
//////		tablaDatos.setTotalWidth(inAncho);
//////		
//////	
//////		//numero item
//////		PdfPCell cell1 = new PdfPCell(new Paragraph("",fontBold));
//////		cell1.setHorizontalAlignment(Element.ALIGN_RIGHT);
//////		cell1.setBorderWidth((float)0.0);
//////		cell1.setBorderColor(new java.awt.Color(166,166,166));
//////		tablaDatos.addCell(cell1);
//////
//////		//numero item
//////	  	cell1 = new PdfPCell(new Paragraph("TOTAL ORDEN DE COMPRA:",fontBold));
//////		cell1.setHorizontalAlignment(Element.ALIGN_LEFT);
//////		cell1.setBorderWidth((float)0.5);
//////		cell1.setBorderColor(new java.awt.Color(166,166,166));
//////		cell1.setBackgroundColor(new Color(217,217,217));
//////		tablaDatos.addCell(cell1);
//////		
//////		//valor total OC
//////		Double valorTotalOC = valorTotal;
//////		String valorTotalOCImpreso = formateador.format(valorTotalOC);
//////		if(Boolean.valueOf(tmpOrdenCompra.impresionDolares)){
//////			valorTotalOC = valorTotalOC/Double.parseDouble(tmpOrdenCompra.tasaCambio);
//////			valorTotalOCImpreso = formateadorDolares.format(valorTotalOC);
//////		}	
//////		
//////		//nombre item
//////		cell1 = new PdfPCell(new Paragraph(valorTotalOCImpreso,fontBold));
//////		cell1.setBorderWidth((float)0.5);
//////		cell1.setHorizontalAlignment(Element.ALIGN_RIGHT);
//////		cell1.setBorderColor(new java.awt.Color(166,166,166));
//////		tablaDatos.addCell(cell1);
			
		
		return tablaDatos;
	}
	

	




	/**
	 * crea la tabla de detalles de asociaciones (aplica para MateriaPrima, Insumos, Muestras, MaquinaExt)
	 *
	 * @param inTablaMateriaPrima --tabla que se imprimira
	 * @param inTitulo --titulo de la tabla
	 * @param inAnchos --anchos de las columnas
	 * 
	 * @return -retorna la tabla creada con la informacion de la tabla recibida
	 *
	 */
	public PdfPTable crearTablaMembrete(Document miDocumento) throws Exception {
    	
    	
    	//encabezado (membrete)
		com.lowagie.text.Image logotipo = com.lowagie.text.Image.getInstance(imagenMembrete, null);
		
		logotipo.scaleToFit(160,70);
		logotipo.setAlignment(com.lowagie.text.Image.DEFAULT); 
		
//		logotipo.setAbsolutePosition(30, miDocumento.getPageSize().height()-100);
		
		PdfPTable encabezado = new PdfPTable(2);
		
		encabezado.setWidthPercentage(100);
		
		int[] anchosMembrete = {35,65};
		encabezado.setWidths(anchosMembrete);

		PdfPCell cell1 = new PdfPCell(logotipo);
		cell1.setHorizontalAlignment(Element.ALIGN_CENTER);
		cell1.setVerticalAlignment(Element.ALIGN_MIDDLE);
		
		cell1.setBorderWidth((float)0.5);
		cell1.setBorderColor(new java.awt.Color(217,217,217));

//		cell1.setFixedHeight(100);
//		cell1.setBorderWidth(1);
	
		encabezado.addCell(cell1);
		
		//celda cotizacion
		PdfPCell cellTitulo = new PdfPCell(new Paragraph("CERTIFICADO LABORAL",fontBoldTitulo));
		cellTitulo.setHorizontalAlignment(Element.ALIGN_CENTER);
		cellTitulo.setVerticalAlignment(Element.ALIGN_MIDDLE);
		cellTitulo.setBorderWidth((float)0.5);
		cellTitulo.setBorderColor(new java.awt.Color(217,217,217));
		
		encabezado.addCell(cellTitulo);
		
//////		PdfPTable tablaInterna = new PdfPTable(1);
//////		
//////		
//////		tablaInterna.addCell(new PdfPCell(new Paragraph("Código: PT-FO-01",fontNormalSmall)));
//////		tablaInterna.addCell(new PdfPCell(new Paragraph("Versión: 01",fontNormalSmall)));
//////		tablaInterna.addCell(new PdfPCell(new Paragraph("Vigencia: 24/03/2022",fontNormalSmall)));
		
//////		PdfPCell cellTablaInterna = new PdfPCell(tablaInterna);
//////		cellTablaInterna.setBorderWidth((float)0.0);
		
		
//////		encabezado.addCell(cellTablaInterna);
		
		encabezado.setTotalWidth(miDocumento.getPageSize().width() - miDocumento.leftMargin() - miDocumento.rightMargin());
       
        return encabezado;
	}
	/**
	 * evento de cambio de pagina (encargado de colocar el encabezado)
	 *
	 * @param writer --escritor utilizado
	 * @param miDocumento --documento que se esta escribiendo
	 *
	 */
	public void onEndPage(PdfWriter writer, Document miDocumento) {
        try{
        	
//        	//Coloca la marca de agua del programa
//			InformeEstandar.establecerMarca(miDocumento);
//			//
			
        	Color colorLineas = new Color(0,0,0);
        	Color color1 = new Color(0,0,0);
        	Color color2 = new Color(0,0,0);
        	Color color3 = new Color(0,0,0);
        	Color color4 = new Color(0,0,0);
        	int anchoBorde = 0;
        	
        	String textoPagina = "Página "+(miDocumento.getPageNumber());      	
        	
			PdfPTable numeroPagina = new PdfPTable(1);
			com.lowagie.text.Font fuenteFecha = com.lowagie.text.FontFactory.getFont("Arial", 9, com.lowagie.text.Font.NORMAL,Color.BLACK);
			PdfPCell celdaPagina = new PdfPCell(new Paragraph(textoPagina,fuenteFecha));
			celdaPagina.setHorizontalAlignment(Element.ALIGN_RIGHT);
			celdaPagina.setBorderWidth(anchoBorde);
			
			numeroPagina.addCell(celdaPagina);
			
			
			if(miDocumento.getPageNumber()!=0){
				crearTablaMembrete(miDocumento).writeSelectedRows(0, -1, miDocumento.leftMargin(), miDocumento.getPageSize().height() - 20,
             								writer.getDirectContent());
			}
			
			
            								
			
			
//			piePagina.setTotalWidth(miDocumento.getPageSize().width() - miDocumento.leftMargin() - miDocumento.rightMargin());
//			
//			piePagina.writeSelectedRows(0, -1, miDocumento.leftMargin(), 45,
//             								writer.getDirectContent());
            
            numeroPagina.setTotalWidth(miDocumento.getPageSize().width() - miDocumento.leftMargin() - miDocumento.rightMargin());
//            numeroPagina.writeSelectedRows(0, -1, miDocumento.leftMargin(), miDocumento.getPageSize().height() - 95,
//             								writer.getDirectContent());
             								
            numeroPagina.writeSelectedRows(0, -1, miDocumento.leftMargin(),  45,
             								writer.getDirectContent()); 								
        }catch(Exception e){
        	e.printStackTrace();
        }
	}
};