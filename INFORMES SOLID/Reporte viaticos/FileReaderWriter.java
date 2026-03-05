/*
 Proyecto Sistema Integrado de Produccion (SOLID) THC
 
 Clase 				FileReaderWriter
 Archivo			FileReaderWriter.java
  
 Fecha Creación 	1 de Marzo de 2008
 Autor  			Marlon Ramirez Jaime 1.130.637.509
 email 				marlonrj@gmail.com
*/

import java.util.*;
import javax.swing.*;
import java.awt.*;
import java.io.*;
import java.text.SimpleDateFormat;

/**
 * Esta clase permite leer y escribir los archivos de parametrizacion del sistema, ademas 
 * escribe en la base de datos los parametros de configuracion globales
 *
 */
public class FileReaderWriter{
	
	/** interface con el manejador de base de datos */
	public ConexionDatos miConexionDatos = null;
	
	//nombre de archivos para configuracion del sistema y persistencia de datos (simple)
	/** guarda la ruta y el nombre del archivo donde se guarda la informacion de la ultima ejecucion */
	public static String nombreArchivoUltimaEjecucion = "system\\test2.sip";
	
	/** guarda la ruta y el nombre del archivo donde se guarda la ultima apariencia utilizada */
	public static String archivoApariencia = "system\\apariencia.sip";
	
	/** guarda la ruta y el nombre del archivo donde se guardan las apariencias disponibles */
	public static String archivoListaDeLooks = "system\\looks.txt";
	
	/** guarda la ruta y el nombre del archivo donde se guarda la informacion de los modulos */
	public static String archivoListaDeModulos = "system\\modulos.txt";
	
	/** guarda la ruta y el nombre del archivo donde se guarda la informacion de los tipos de modulos */
	public static String archivoListaDeTipoModulos = "system\\tipos modulos.txt";	
		
	//nombres de archivos de parametros para generacion de pfd
	/** guarda la ruta y el nombre del archivo de parametros de conexion */
	public static String archivoParametrosConexion = "system\\parametrosConexion.sip";
	
	/** guarda la ruta y el nombre del archivo de parametros de informe */
	public static String archivoParametrosInforme = "system\\parametrosInforme.sip"; //usado para la direccion del lector de pdf
	
	/** guarda la ruta y el nombre del archivo con el password de configuracion critica */
	public static String archivoPassword = "system\\pass.sip";
	
	/** guarda la ruta y el nombre del archivo con el esquema de la base de datos */
	public static String archivoScriptDB = "system\\scriptDB.sip";
	
	/** guarda la ruta y el nombre del archivo con los triggers de la base de datos */
	public static String archivoScriptTriggers = "system\\scriptTDB.sip";
	
	/** guarda la ruta del directorio donde se almacenan las imagenes */
	public static String folderImages = "images\\";
	
	/** guarda la ruta y el nombre del archivo de ayuda */
	public static String archivoAyuda = "ayuda\\ayudaSIP.chm";
	
	/** guarda la ruta y el nombre del programa visualizador de la ayuda */
	public static String visualizadorAyuda = "hh.exe";
	
	/** guarda la ruta del directorio de archivos temporales del sip */
	public static String directorioTemporal = "tmp\\";
	
	///////////////////////////variables de conexion a la base de datos//////////////////////////////
	/////////////////////////////////////////////////////////////////////////////////////////////////
	/** guarda el nombre de la base de datos que se configurara */
	public String nombreBaseDatos = "";
	
	/** guarda la direccion de la base de datos que se configurara */
	public String direccion = "";
	
	/** guarda el puerto de conexion del servidor que tiene a base de datos */
	public String puerto = "";
	
	/** guarda el nombre de usuario que se utilizara para acceder a la base de datos */
	public String nombreUsuario = "";
	
	/** guarda el password que se utilizara para acceder a la base de datos */
	public String password = "";
	
	////////////////////variables para informes generales (membretes en informe)/////////////////////
	/////////////////////////////////////////////////////////////////////////////////////////////////
	/** guarda la direccion del programa que se usara como visualizador de los informes generados */
	public String dirVisualizador = "";
	
	/** guarda la la opcion de visualizacion de informes en pdf 0 para guardar 1 para mostrardirectamente*/
	public String opcionGuardarPDF = "";
	
	////////////////////variables para formatos de cotizacion de venta/////////////////////
	/////////////////////////////////////////////////////////////////////////////////////////////////
	/** guarda un identificador de la forma de impresion del identificador 1 o 0 para codigo o referencia */
	public String idCotizacionVenta = "";

	/** guarda la contenido de cotizacion */
	public String contenidoCotizacionVenta = "";
	
	/** guarda la despedida de cotizacion */
	public String despedidaCotizacionVenta = "";
	
	/** guarda el remitente de la cotizacion */
	public String remitenteCotizacionVenta = "";
	
	/** guarda el tamaño de fuente utilizada en los textos de cotizacion venta */
	public String fuenteCotizacionVenta = "";
	
	////////////////////variables para formatos de orden de ejecucion/////////////////////
	/////////////////////////////////////////////////////////////////////////////////////////////////
	/** guarda un identificador de la forma de impresion del identificador 1 o 0 para codigo o referencia */
	public String idOrdenEjecucion = "";
	
	/** guarda la contenido de la orden de ejecucion */
	public String anotacionesOrdenEjecucion = "";
	
	/** guarda la despedida de la orden de ejecucion */
	public String firmaOrdenEjecucion = "";
	
	/** guarda el tamaño de fuente utilizada en las anotaciones de la orden de ejecucion */
	public String fuenteOrdenEjecucion = "";
	
	////////////////////variables para formatos de orden de ejecucion/////////////////////
	/////////////////////////////////////////////////////////////////////////////////////////////////
	
	/** guarda las anotaciones de la OIT */
	public String anotacionesOIT = "";
	
	/** guarda el texto del primer rotulo de firma de OIT */
	public String rotuloFirmaOIT1 = "";
	
	/** guarda el texto del segundo rotulo de firma de OIT */
	public String rotuloFirmaOIT2 = "";
	
	/** guarda el tamaño de fuente utilizada en los textos de OIT */
	public String fuenteOIT = "";
	
	//////////////////variables para configuraciones iniciales de apariencia de SOLID//////////////////
	/////////////////////////////////////////////////////////////////////////////////////////////////
	/** guarda el ultimo look and feel de la aplicacion configurado */
	public String lookAndFeel = "";
	
	/** guarda los nombres de looks ej (Look estilo Java) */
	public Vector vNombresLooks = new Vector();
	
	/** guarda las direcciones de las clases de looks ej (com.sun.java.....) */
	public Vector vDireccionesLooks = new Vector();
	
	////////////////////variables para configuraciones de modulos cargados de SOLID////////////////////
	/////////////////////////////////////////////////////////////////////////////////////////////////
	/** guarda los codigos de los modulos */
	public Vector vCodigosModulos = new Vector(); //vector de codigos de modulos
	
	/** guarda los nombres de los modulos */
	public Vector vNombresModulos = new Vector(); //vector de nombres de modulos
	
	/** guarda los nombres a mostrar de los modulos */
	public Vector vVisualizacionesModulos = new Vector(); //vector de visualizaciones de modulos
	
	/** guarda los codigos de los tipos de modulos */
	public Vector vTiposModulos = new Vector(); //vector de tipos de modulos
	
	/////////////////variables para configuraciones de tipos de modulos cargados de SOLID//////////////
	/////////////////////////////////////////////////////////////////////////////////////////////////
	/** guarda los codigos de los tipos de modulos */
	public Vector vCodigosTipoModulos = new Vector();
	
	/** guarda los nombres a mostrar de los tipos de modulos */
	public Vector vVisualizacionesTipoModulos = new Vector();
	
	/** guarda las descripciones de los tipos de modulos */
	public Vector vDescripcionesTipoModulos = new Vector();
	
	/////////////////variables para almacenamiento de fechas de ultima ejecucion de SOLID//////////////
	/////////////////////////////////////////////////////////////////////////////////////////////////
	/** guarda la fecha de ultima ejecucion del sistema */
	public String fechaUltimaEjecucion = "";
	
	/////////////////variables para almacenamiento del pie de pagina de informes de SOLID//////////////
	/////////////////////////////////////////////////////////////////////////////////////////////////
	/** guarda el pie de pagina que se usara para los informes del sistema */
	public String piePaginaInforme = "";
	
	////////////////////variables para formatos de orden de compra/////////////////////
	/////////////////////////////////////////////////////////////////////////////////////////////////
	/** guarda la contenido de la orden de compra */
	public String anotacionesOrdenCompra = "";
	
	////////////////////variables para formatos de carnet/////////////////////
	/////////////////////////////////////////////////////////////////////////////////////////////////
	/** guarda las anotaciones de la cara posterior del carnet de maquinas internas */
	public String anotacionesCarnetMaquinaInt = "";
	
	/** guarda las anotaciones de la cara posterior del carnet de personal */
	public String anotacionesCarnetPersonal = "";
	
	////////////////////variables para parametros de costeo/////////////////////
	/////////////////////////////////////////////////////////////////////////////////////////////////
	/** guarda el valor del porcentaje de costos indirectos de fabricacion del producto */
	public String porcentajeCostosIndirectosFabricacion = "";
	
	/** guarda el valor del porcentaje de comisiones del producto */
	public String porcentajeComisiones = "";
	
	/** guarda el valor del porcentaje de margen de utilidad del producto */
	public String porcentajeMargenUtilidad = "";
	
	////////////////////variables de tiempos de modulo de registro de OIT/////////////////////
	/////////////////////////////////////////////////////////////////////////////////////////////////
	/** guarda la cantidad de segundos de habilitacion de registro de OIT */
	public String segundosRegistro = "";
	
	/** guarda el cantidad de segundos de presentacion de la informacion de la OIT marcada */
	public String segundosPresentacionOIT = "";
	
	////////////////////variables para passwor de configuracion/////////////////////
	/////////////////////////////////////////////////////////////////////////////////////////////////
	/** guarda el valor del password de configuracion */
	public String passwordConfiguracion = "";
	
	////////////////////variables para impresion de factura/////////////////////
	/////////////////////////////////////////////////////////////////////////////////////////////////
	/** guarda el inicio de la numeracion de facturas */
	public String inicioNumeracion = "";
	
	/** guarda el final de la numeracion de facturas */
	public String finNumeracion = "";
	
	/** guarda el texto legal en facturas */
	public String textoLegalFactura = "";
	
	/** guarda el texto legal de numeracion de facturas */
	public String textoLegalNumeracionFactura = "";
	
	/** guarda el texto de pie de pagina de facturas */
	public String textoPiePaginaFactura = "";

	/** guarda el texto adicional en factura */
	public String textoAdicionalFactura = "";
		
	/** guarda el tamaño de fuente utilizada en el contenido de la factura */
	public String fuenteFactura = "";
	
	/** guarda el tamaño de fuente utilizada en el texto de informacion legal de la empresa */
	public String fuenteFacturaTextoLegal = "";
	
	/** guarda el tamaño de fuente utilizada en el texto de numeracion de facturas */
	public String fuenteFacturaTextoNumeracion = "";
	
	/** guarda el tamaño de fuente utilizada en el pie de pagina de la factura */
	public String fuenteFacturaPiePagina = "";
	
	/** guarda el tamaño de fuente utilizada en el texto adicional en factura */
	public String fuenteFacturaTextoAdicional = "";
	
	/** guarda el prefijo de archivo de facturas */
	public String prefijoArchivoFactura = "";
	
	/** determina la forma de impresion de las facturas 0 -> completa, 1 -> impresion en formato */
	public String idImpresionFactura = "";
	
	////////////////////otras variables/////////////////////
	/////////////////////////////////////////////////////////////////////////////////////////////////
	/** guarda el tamaño de fuente utilizada en la orden de compra */
	public String fuenteOrdenCompra = "";
	
	/** guarda el tamaño de fuente utilizada en la requisicion */
	public String fuenteRequisicion = "";
	
	/** dia del mes utilizado para el corte en informes de gerencia */
	public String diaCorteInformes = "";
	
	/** frecuencia de actualizacion informes de gerencia */
	public String frecuenciaActualizacionInformes = "";
	
	////////////////////variables para impresion de cuentas de cobro/////////////////////
	/////////////////////////////////////////////////////////////////////////////////////////////////
	
	/** guarda el texto legal en CC */
	public String textoLegalCC = "";
	
	/** guarda el texto de pie de pagina de CC */
	public String textoPiePaginaCC = "";

	/** guarda el texto adicional en factura */
	public String textoAdicionalCC = "";
		
	/** guarda el tamaño de fuente utilizada en el contenido de la CC */
	public String fuenteCC = "";
	
	/** guarda el tamaño de fuente utilizada en el texto de informacion legal de la empresa */
	public String fuenteCCTextoLegal = "";
	
	/** guarda el tamaño de fuente utilizada en el pie de pagina de la CC */
	public String fuenteCCPiePagina = "";
	
	/** guarda el tamaño de fuente utilizada en el texto adicional en CC */
	public String fuenteCCTextoAdicional = "";
	
	//////
	
	/** guarda el numero de dias en mora antes de bloquear a un cliente */
	public String numeroDiasBloqueo = "";
		
	
	/** 
	 * constructor por defecto
	 * 
	 * @param inConexionDatos -interfaz con base de datos que se utilizara
	 *
	 */
	public FileReaderWriter(ConexionDatos inConexionDatos){
		miConexionDatos = inConexionDatos;
	}
	
//	/** 
//	 * constructor por defecto
//	 */
//	public FileReaderWriter(){
//	
//	}
	
	/////////////////metodos de carga de variables de conexion a la base de datos////////////////////
	/////////////////////////////////////////////////////////////////////////////////////////////////
	/**
	 * carga el ultimo look and feel de la aplicacion configurado desde un archivo (archivoApariencia)
	 *
	 */
	public void cargarLookAndFeel(){
		try{
    		ObjectInputStream in = new ObjectInputStream(new FileInputStream(archivoApariencia));
            //carga los parametros
            lookAndFeel = (String)in.readObject();
            
            in.close();
       	} catch (ClassNotFoundException f) {
       	    System.out.println(f);
       	} catch (IOException f) {
       	    System.out.println(f);
       	    JOptionPane.showMessageDialog(null,"El archivo de configuracion no se encuentra");
       	}
	}
	/**
	 * guarda el ultimo look and feel de la aplicacion configurado en un archivo (archivoApariencia)
	 *
	 */
	public void guardarLookAndFeel(){
		try{
     		ObjectOutputStream out = new ObjectOutputStream(new FileOutputStream(archivoApariencia));
     		
			//escribe los parametros
			out.writeObject(lookAndFeel);
			
			//
    		out.close();
    	}catch (IOException f) {
        	System.out.println(f);
        	JOptionPane.showMessageDialog(null,"Ha ocurrido un error en la modificacion del archivo de apariencia");
    	}
	}
	/**
	 * carga los nombres y direcciones de looksAndFeels de un archivo (archivoListaDeLooks), los coloca en 2 vectores
	 * en el archivo primero van las direcciones looks y luego los nombres en cada fila
	 *
	 */
	public void cargarLooks(){
		try{
			FileInputStream archivo = new FileInputStream(archivoListaDeLooks);
    		BufferedReader entrada = new BufferedReader(new InputStreamReader(archivo));
    	
	    	String linea = ""; 
	    	while((linea = entrada.readLine())!=null){
	    		StringTokenizer miTK = new StringTokenizer(linea,",");
	    		
				vDireccionesLooks.add(miTK.nextToken().trim());
				vNombresLooks.add(miTK.nextToken().trim());
	    	}
		}catch(Exception e){
			e.printStackTrace();
		}
	}
////////	/**
////////	 * carga los codigos, nombres, visualizacion y tipo de modulo respectivamente de un archivo (archivoListaDeModulos), los coloca en 4 vectores
////////	 *
////////	 */
////////	public void cargarInformacionModulos(){
////////		try{
////////			FileInputStream archivo = new FileInputStream(archivoListaDeModulos);
////////    		BufferedReader entrada = new BufferedReader(new InputStreamReader(archivo));
////////    	
////////	    	String linea = ""; 
////////	    	while((linea = entrada.readLine())!=null){
////////	    		StringTokenizer miTK = new StringTokenizer(linea,",");
////////	    		
////////				vCodigosModulos.add(miTK.nextToken().trim());
////////				vNombresModulos.add(miTK.nextToken().trim());
////////				vVisualizacionesModulos.add(miTK.nextToken().trim());
////////				vTiposModulos.add(miTK.nextToken().trim());
////////	    	}
////////		}catch(Exception e){
////////			e.printStackTrace();
////////		}
////////		//impresion de control
//////////		Restriccion.imprimirVector(vCodigosModulos);
//////////		Restriccion.imprimirVector(vNombresModulos);
//////////		Restriccion.imprimirVector(vVisualizacionesModulos);
//////////		Restriccion.imprimirVector(vTiposModulos);
////////	}
////////	/**
////////	 * carga los codigos, nombres a mostrar y descripciones de los tipos de modulos respectivamente de un archivo (archivoListaDeTipoModulos),
////////	 * los coloca en 3 vectores
////////	 *
////////	 */
////////	public void cargarInformacionTiposModulos(){
////////		try{
////////			FileInputStream archivo = new FileInputStream(archivoListaDeTipoModulos);
////////    		BufferedReader entrada = new BufferedReader(new InputStreamReader(archivo));
////////    	
////////	    	String linea = ""; 
////////	    	while((linea = entrada.readLine())!=null){
////////	    		StringTokenizer miTK = new StringTokenizer(linea,",");
////////	    		
////////				vCodigosTipoModulos.add(miTK.nextToken().trim());
////////				vVisualizacionesTipoModulos.add(miTK.nextToken().trim());
////////				vDescripcionesTipoModulos.add(miTK.nextToken().trim());
////////	    	}
////////		}catch(Exception e){
////////			e.printStackTrace();
////////		}
////////		//impresion de control
//////////		Restriccion.imprimirVector(vCodigosTipoModulos);
//////////		Restriccion.imprimirVector(vVisualizacionesTipoModulos);
//////////		Restriccion.imprimirVector(vDescripcionesTipoModulos);
////////	}
	/**
	 * carga los codigos, nombres, visualizacion y tipo de modulo respectivamente de un archivo (archivoListaDeModulos), los coloca en 4 vectores
	 *
	 */
	public void cargarInformacionModulos(){
		
		vCodigosModulos.add("12"); vNombresModulos.add("personal"); vVisualizacionesModulos.add("Empleados");vTiposModulos.add("1");
		vCodigosModulos.add("14"); vNombresModulos.add("actividadesUsuarios"); vVisualizacionesModulos.add("Actividades de usuarios");vTiposModulos.add("5");
		vCodigosModulos.add("6"); vNombresModulos.add("opciones"); vVisualizacionesModulos.add("Parametrización general");vTiposModulos.add("6");
		vCodigosModulos.add("7"); vNombresModulos.add("configuracion"); vVisualizacionesModulos.add("Parametrización de operación");vTiposModulos.add("6");
		vCodigosModulos.add("8"); vNombresModulos.add("usuarios"); vVisualizacionesModulos.add("Administración de usuarios");vTiposModulos.add("6");
		vCodigosModulos.add("401"); vNombresModulos.add("Consignaciones"); vVisualizacionesModulos.add("Consignaciones");vTiposModulos.add("2");
		vCodigosModulos.add("402"); vNombresModulos.add("Legalizaciones"); vVisualizacionesModulos.add("Legalizaciones");vTiposModulos.add("2");
		vCodigosModulos.add("403"); vNombresModulos.add("Viaticantes"); vVisualizacionesModulos.add("Viaticantes");vTiposModulos.add("2");
		vCodigosModulos.add("406"); vNombresModulos.add("Catalogo de Productos"); vVisualizacionesModulos.add("Catálogo de Productos");vTiposModulos.add("3");
		vCodigosModulos.add("404"); vNombresModulos.add("Solicitudes de Materiales"); vVisualizacionesModulos.add("Solicitudes de Materiales");vTiposModulos.add("3");
		vCodigosModulos.add("405"); vNombresModulos.add("BDG Costos"); vVisualizacionesModulos.add("BDG Costos");vTiposModulos.add("4");
		

////////		try{
////////			FileInputStream archivo = new FileInputStream(archivoListaDeModulos);
////////    		BufferedReader entrada = new BufferedReader(new InputStreamReader(archivo));
////////    	
////////	    	String linea = ""; 
////////	    	while((linea = entrada.readLine())!=null){
////////	    		StringTokenizer miTK = new StringTokenizer(linea,",");
////////	    		
////////				vCodigosModulos.add(miTK.nextToken().trim());
////////				vNombresModulos.add(miTK.nextToken().trim());
////////				vVisualizacionesModulos.add(miTK.nextToken().trim());
////////				vTiposModulos.add(miTK.nextToken().trim());
////////	    	}
////////		}catch(Exception e){
////////			e.printStackTrace();
////////		}
////////		//impresion de control
//////////		Restriccion.imprimirVector(vCodigosModulos);
//////////		Restriccion.imprimirVector(vNombresModulos);
//////////		Restriccion.imprimirVector(vVisualizacionesModulos);
//////////		Restriccion.imprimirVector(vTiposModulos);
	}
	/**
	 * carga los codigos, nombres a mostrar y descripciones de los tipos de modulos respectivamente de un archivo (archivoListaDeTipoModulos),
	 * los coloca en 3 vectores
	 *
	 */
	public void cargarInformacionTiposModulos(){
		
		
		vCodigosTipoModulos.add("1"); vVisualizacionesTipoModulos.add("Gestion Humana"); vDescripcionesTipoModulos.add("menu de Gestion Humana");
		vCodigosTipoModulos.add("2"); vVisualizacionesTipoModulos.add("Viaticos"); vDescripcionesTipoModulos.add("menu de viaticos");
		vCodigosTipoModulos.add("3"); vVisualizacionesTipoModulos.add("Logistica"); vDescripcionesTipoModulos.add("menu de logistica");
		vCodigosTipoModulos.add("4"); vVisualizacionesTipoModulos.add("Costos"); vDescripcionesTipoModulos.add("menu de costos");
		vCodigosTipoModulos.add("5"); vVisualizacionesTipoModulos.add("Informes"); vDescripcionesTipoModulos.add("menu de informes");
		vCodigosTipoModulos.add("6"); vVisualizacionesTipoModulos.add("Parametrizacion"); vDescripcionesTipoModulos.add("menu de parametrización");
		
		
		
////////		try{
////////			FileInputStream archivo = new FileInputStream(archivoListaDeTipoModulos);
////////    		BufferedReader entrada = new BufferedReader(new InputStreamReader(archivo));
////////    	
////////	    	String linea = ""; 
////////	    	while((linea = entrada.readLine())!=null){
////////	    		StringTokenizer miTK = new StringTokenizer(linea,",");
////////	    		
////////				vCodigosTipoModulos.add(miTK.nextToken().trim());
////////				vVisualizacionesTipoModulos.add(miTK.nextToken().trim());
////////				vDescripcionesTipoModulos.add(miTK.nextToken().trim());
////////	    	}
////////		}catch(Exception e){
////////			e.printStackTrace();
////////		}
////////		//impresion de control
//////////		Restriccion.imprimirVector(vCodigosTipoModulos);
//////////		Restriccion.imprimirVector(vVisualizacionesTipoModulos);
//////////		Restriccion.imprimirVector(vDescripcionesTipoModulos);
	}
	/**
	 * carga la fecha de la ultima ejecucion de la aplicacion desde un archivo (nombreArchivoUltimaEjecucion)
	 *
	 */
	public void cargarFechaUltimaEjecucion(){
		try{
    		ObjectInputStream in = new ObjectInputStream(new FileInputStream(nombreArchivoUltimaEjecucion));
            //carga los parametros
            fechaUltimaEjecucion = (String)in.readObject();
            
            in.close();
       	} catch (ClassNotFoundException f) {
       	    System.out.println(f);
       	} catch (IOException f) {
       	    System.out.println(f);
       	    JOptionPane.showMessageDialog(null,"El archivo de configuracion no se encuentra");
       	}
	}
	/**
	 * guarda la fecha de la ultima ejecucion de la aplicacion en un archivo (nombreArchivoUltimaEjecucion)
	 *
	 */
	public void guardarFechaUltimaEjecucion(){
		try{
     		ObjectOutputStream out = new ObjectOutputStream(new FileOutputStream(nombreArchivoUltimaEjecucion));
     		
			//escribe los parametros
			out.writeObject(fechaUltimaEjecucion);
			
			//
    		out.close();
    	}catch (IOException f) {
        	System.out.println(f);
        	JOptionPane.showMessageDialog(null,"Ha ocurrido un error en la modificacion del archivo de fecha");
    	}
	}
	/**
	 * guarda los parametros de conexion en un archivo (achivoParametrosConexion) y prueba la configuracion
	 *
	 */
	public void guardarParametrosConexion(ConexionDatos inConexionDatos){
		int opcion = 0;
		try{
			ConexionDatos tmpConexionDatos = new ConexionDatos();
			if(tmpConexionDatos.conexionPostgreSQL(direccion+":"+puerto, nombreBaseDatos, nombreUsuario, password)==0){
				opcion = JOptionPane.showConfirmDialog(null,"La conexion no se pudo establecer, desea continuar","Confirmacion",0);
			}
			tmpConexionDatos.terminarConexion();
			if(opcion == 0){
				//
				//cierra sesion
				//establece los parametros
				//
				
                inConexionDatos.terminarConexion();
                inConexionDatos.conexionPostgreSQL(direccion+":"+puerto,nombreBaseDatos,nombreUsuario,password);
         		ObjectOutputStream out = new ObjectOutputStream(new FileOutputStream(archivoParametrosConexion));
				//escribe los parametros
				out.writeObject(nombreBaseDatos);
				out.writeObject(direccion);
				out.writeObject(puerto);
				out.writeObject(nombreUsuario);
				out.writeObject(password);
				
				JOptionPane.showMessageDialog(null,"Se han cambiado los parametros con exito");
        		out.close();
			}
    	}catch (IOException f) {
        	System.out.println(f);
        	JOptionPane.showMessageDialog(null,"Ha ocurrido un error en la modificacion del archivo de parametros");
    	}
	}
	
	/**
	 * guarda los parametros de conexion en un archivo (achivoParametrosConexion) y prueba la configuracion y conecta 
	 * con los nuevos parametros a la base 
	 *
	 */
	public void guardarParametrosConexion2(ConexionDatos inConexionDatos, SOLID framePadre){
		int opcion = 0;
		try{
			ConexionDatos tmpConexionDatos = new ConexionDatos();
			if(tmpConexionDatos.conexionPostgreSQL(direccion+":"+puerto, nombreBaseDatos, nombreUsuario, password)==0){
				opcion = JOptionPane.showConfirmDialog(null,"La conexion no se pudo establecer, desea continuar","Confirmacion",0);
			}
			tmpConexionDatos.terminarConexion();
			if(opcion == 0){
				//
				//cierra sesion
				//establece los parametros
				framePadre.cerrarSesion();
				//
				
                inConexionDatos.terminarConexion();
                inConexionDatos.conexionPostgreSQL(direccion+":"+puerto,nombreBaseDatos,nombreUsuario,password);
         		ObjectOutputStream out = new ObjectOutputStream(new FileOutputStream(archivoParametrosConexion));
				//escribe los parametros
				out.writeObject(nombreBaseDatos);
				out.writeObject(direccion);
				out.writeObject(puerto);
				out.writeObject(nombreUsuario);
				out.writeObject(password);
				
				JOptionPane.showMessageDialog(null,"Se han cambiado los parametros con exito");
        		out.close();
			}
    	}catch (IOException f) {
        	System.out.println(f);
        	JOptionPane.showMessageDialog(null,"Ha ocurrido un error en la modificacion del archivo de parametros");
    	}
	}
	/**
	 * carga los parametros de conexion desde un archivo (achivoParametrosConexion)
	 *
	 */
	public void cargarParametrosConexion(){
		try{
    		ObjectInputStream in = new ObjectInputStream(new FileInputStream(archivoParametrosConexion));
            //carga los parametros
            nombreBaseDatos = (String)in.readObject();
            direccion = (String)in.readObject();
            puerto = (String)in.readObject();
            nombreUsuario = (String)in.readObject();
            password = (String)in.readObject();
            
            in.close();
       	} catch (ClassNotFoundException f) {
       	    System.out.println(f);
       	} catch (IOException f) {
       	    System.out.println(f);
       	    JOptionPane.showMessageDialog(null,"El archivo de configuracion no se encuentra, debe configurar\n"+
       	    	"los parametros de conexion de nuevo");
       	}
	}
	
	/**
	 * guarda los parametros de informe en un archivo (achivoParametrosInforme)
	 *
	 */
	public void guardarParametrosInforme(){
		int numeroActualizaciones = 0;
		try{
			ObjectOutputStream out = new ObjectOutputStream(new FileOutputStream(archivoParametrosInforme));
			//escribe los parametros
			out.writeObject(dirVisualizador);
			out.writeObject(opcionGuardarPDF);
//			out.writeObject(dirVisualizador.replaceAll("\\\\","/"));
			
			
    		out.close();
//			JOptionPane.showMessageDialog(null,"Se han cambiado los parametros de informes con exito");
    	}catch (IOException f) {
        	System.out.println(f);
        	numeroActualizaciones+=1;
//        	JOptionPane.showMessageDialog(null,"Ha ocurrido un error en la modificacion del archivo de parametros de informes");
    	}
    	ParametroInforme tmpParametroInforme = new ParametroInforme(
																	"38",
																	diaCorteInformes,
																	miConexionDatos
																	);
		numeroActualizaciones += tmpParametroInforme.actualizar();
		
		//
		tmpParametroInforme = new ParametroInforme(
													"39",
													frecuenciaActualizacionInformes,
													miConexionDatos
													);
		numeroActualizaciones += tmpParametroInforme.actualizar();
		
		if(numeroActualizaciones == 2){
			JOptionPane.showMessageDialog(null,"Se han cambiado los parametros de informes con exito");
		}else{
			JOptionPane.showMessageDialog(null,"Ha ocurrido un error en la modificacion del archivo de parametros de informes");
		}
	}
	
	/**
	 * carga los parametros de informe desde un archivo (achivoParametrosInforme)
	 *
	 */
	public void cargarParametrosInforme(){
		ParametroInforme tmpParametroInforme = new ParametroInforme(
																null,
																miConexionDatos
																);
        
        tmpParametroInforme.cargar("codigo", "38");
        diaCorteInformes = tmpParametroInforme.parametro;
        
        //
        tmpParametroInforme.cargar("codigo", "39");
        frecuenciaActualizacionInformes = tmpParametroInforme.parametro;
        
		try{
    		ObjectInputStream in = new ObjectInputStream(new FileInputStream(archivoParametrosInforme));
            //carga los parametros
            dirVisualizador = (String)in.readObject();
            opcionGuardarPDF = (String)in.readObject();
//            dirVisualizador = ((String)in.readObject()).replaceAll("/","\\\\");
                        
            in.close();
       	} catch (ClassNotFoundException f) {
       	    System.out.println(f);
       	} catch (IOException f) {
       	    System.out.println(f);
       	    JOptionPane.showMessageDialog(null,"El archivo de configuracion de informes no se encuentra, debe configurar\n"+
       	    	"los parametros de conexion de nuevo");
       	}
	}
	
	/**
	 * guarda los parametros de cotizacionVenta (archivoParametrosCotizacionVenta)
	 *
	 */
	public void guardarParametrosCotizacionVenta(){
		int numeroActualizaciones = 0;
		
		ParametroInforme tmpParametroInforme = new ParametroInforme(
																	"2",
																	idCotizacionVenta,
																	miConexionDatos
																	);
		numeroActualizaciones += tmpParametroInforme.actualizar();
		
		//
		tmpParametroInforme = new ParametroInforme(
													"3",
													fuenteCotizacionVenta,
													miConexionDatos
													);
		numeroActualizaciones += tmpParametroInforme.actualizar();
			
		//
		tmpParametroInforme = new ParametroInforme(
													"4",
													contenidoCotizacionVenta,
													miConexionDatos
													);
		numeroActualizaciones += tmpParametroInforme.actualizar();
		
		//
		tmpParametroInforme = new ParametroInforme(
													"5",
													despedidaCotizacionVenta,
													miConexionDatos
													);
		numeroActualizaciones += tmpParametroInforme.actualizar();
		
		//
		tmpParametroInforme = new ParametroInforme(
													"6",
													remitenteCotizacionVenta,
													miConexionDatos
													);
		numeroActualizaciones += tmpParametroInforme.actualizar();
		
		//
		if(numeroActualizaciones == 5){
			JOptionPane.showMessageDialog(null,"Se han cambiado los parametros de cotizacion con exito");
		}else{
			JOptionPane.showMessageDialog(null,"Ha ocurrido un error en la modificacion del archivo de parametros de cotizacion");
		}
	}
	
	/**
	 * carga los parametros de cotizacionVenta desde un archivo (archivoParametrosCotizacionVenta)
	 *
	 */
	public void cargarParametrosCotizacionVenta(){
		ParametroInforme tmpParametroInforme = new ParametroInforme(
																null,
																miConexionDatos
																);
        
        tmpParametroInforme.cargar("codigo", "2");
        idCotizacionVenta = tmpParametroInforme.parametro;
        
        //
        tmpParametroInforme.cargar("codigo", "3");
        fuenteCotizacionVenta = tmpParametroInforme.parametro;
        
        //
        tmpParametroInforme.cargar("codigo", "4");
        contenidoCotizacionVenta = tmpParametroInforme.parametro;
        
        //
        tmpParametroInforme.cargar("codigo", "5");
        despedidaCotizacionVenta = tmpParametroInforme.parametro;
        
        //
        tmpParametroInforme.cargar("codigo", "6");
        remitenteCotizacionVenta = tmpParametroInforme.parametro;
	}
	
	/**
	 * guarda los parametros de OrdenEjecucion (archivoParametrosOrdenEjecucion)
	 *
	 */
	public void guardarParametrosOrdenEjecucion(){
		int numeroActualizaciones = 0;
		
		ParametroInforme tmpParametroInforme = new ParametroInforme(
																	"7",
																	idOrdenEjecucion,
																	miConexionDatos
																	);
		numeroActualizaciones += tmpParametroInforme.actualizar();
		
		//
		tmpParametroInforme = new ParametroInforme(
													"8",
													fuenteOrdenEjecucion,
													miConexionDatos
													);
		numeroActualizaciones += tmpParametroInforme.actualizar();
			
		//
		tmpParametroInforme = new ParametroInforme(
													"9",
													anotacionesOrdenEjecucion,
													miConexionDatos
													);
		numeroActualizaciones += tmpParametroInforme.actualizar();
		
		//
		tmpParametroInforme = new ParametroInforme(
													"10",
													firmaOrdenEjecucion,
													miConexionDatos
													);
		numeroActualizaciones += tmpParametroInforme.actualizar();
		
		//
		if(numeroActualizaciones == 4){
			JOptionPane.showMessageDialog(null,"Se han cambiado los parametros de orden de ejecucion con exito");
		}else{
			JOptionPane.showMessageDialog(null,"Ha ocurrido un error en la modificacion del archivo de orden de ejecucion");
		}
	}
	
	/**
	 * carga los parametros de OrdenEjecucion desde un archivo (archivoParametrosOrdenEjecucion)
	 *
	 */
	public void cargarParametrosOrdenEjecucion(){
		ParametroInforme tmpParametroInforme = new ParametroInforme(
																null,
																miConexionDatos
																);
        
        tmpParametroInforme.cargar("codigo", "7");
        idOrdenEjecucion = tmpParametroInforme.parametro;
        
        //
        tmpParametroInforme.cargar("codigo", "8");
        fuenteOrdenEjecucion = tmpParametroInforme.parametro;
        
        //
        tmpParametroInforme.cargar("codigo", "9");
        anotacionesOrdenEjecucion = tmpParametroInforme.parametro;
        
        //
        tmpParametroInforme.cargar("codigo", "10");
        firmaOrdenEjecucion = tmpParametroInforme.parametro;
	}
	
	/**
	 * guarda los parametros de OIT (archivoParametrosOIT)
	 *
	 */
	public void guardarParametrosOIT(){
		int numeroActualizaciones = 0;
		
		ParametroInforme tmpParametroInforme = new ParametroInforme(
																	"11",
																	rotuloFirmaOIT1,
																	miConexionDatos
																	);
		numeroActualizaciones += tmpParametroInforme.actualizar();
		
		//
		tmpParametroInforme = new ParametroInforme(
													"12",
													rotuloFirmaOIT2,
													miConexionDatos
													);
		numeroActualizaciones += tmpParametroInforme.actualizar();
			
		//
		tmpParametroInforme = new ParametroInforme(
													"13",
													fuenteOIT,
													miConexionDatos
													);
		numeroActualizaciones += tmpParametroInforme.actualizar();
		
		//
		tmpParametroInforme = new ParametroInforme(
													"14",
													anotacionesOIT,
													miConexionDatos
													);
		numeroActualizaciones += tmpParametroInforme.actualizar();
		
		//
		if(numeroActualizaciones == 4){
			JOptionPane.showMessageDialog(null,"Se han cambiado los parametros de orden interna de trabajo con exito");
		}else{
			JOptionPane.showMessageDialog(null,"Ha ocurrido un error en la modificacion del archivo de parametros de orden interna de trabajo");
		}
	}
	
	/**
	 * carga los parametros de OrdenEjecucion desde un archivo (archivoParametrosOIT)
	 *
	 */
	public void cargarParametrosOIT(){
		ParametroInforme tmpParametroInforme = new ParametroInforme(
																null,
																miConexionDatos
																);
        
        tmpParametroInforme.cargar("codigo", "11");
        rotuloFirmaOIT1 = tmpParametroInforme.parametro;
        
        //
        tmpParametroInforme.cargar("codigo", "12");
        rotuloFirmaOIT2 = tmpParametroInforme.parametro;
        
        //
        tmpParametroInforme.cargar("codigo", "13");
        fuenteOIT = tmpParametroInforme.parametro;
        
        //
        tmpParametroInforme.cargar("codigo", "14");
        anotacionesOIT = tmpParametroInforme.parametro;
	}
	/**
	 * guarda los parametros de orden de compra (archivoParametrosOIT)
	 *
	 */
	public void guardarParametrosOrdenCompra(){
		int numeroActualizaciones = 0;
		
		ParametroInforme tmpParametroInforme = new ParametroInforme(
																	"15",
																	anotacionesOrdenCompra,
																	miConexionDatos
																	);
		numeroActualizaciones += tmpParametroInforme.actualizar();
		
		tmpParametroInforme = new ParametroInforme(
																	"36",
																	fuenteOrdenCompra,
																	miConexionDatos
																	);
		numeroActualizaciones += tmpParametroInforme.actualizar();
		
		tmpParametroInforme = new ParametroInforme(
																	"37",
																	fuenteRequisicion,
																	miConexionDatos
																	);
		numeroActualizaciones += tmpParametroInforme.actualizar();
		
			
		//
		if(numeroActualizaciones == 3){
			JOptionPane.showMessageDialog(null,"Se han cambiado los parametros de orden de compra con exito");
		}else{
			JOptionPane.showMessageDialog(null,"Ha ocurrido un error en la modificacion del archivo de parametros de orden de compra");
		}
	}
	
	/**
	 * carga los parametros de orden de compra desde un archivo (archivoParametrosOIT)
	 *
	 */
	public void cargarParametrosOrdenCompra(){
		ParametroInforme tmpParametroInforme = new ParametroInforme(
																null,
																miConexionDatos
																);
        
        tmpParametroInforme.cargar("codigo", "15");
        anotacionesOrdenCompra = tmpParametroInforme.parametro;
        
        tmpParametroInforme.cargar("codigo", "36");
        fuenteOrdenCompra = tmpParametroInforme.parametro;
        
        tmpParametroInforme.cargar("codigo", "37");
        fuenteRequisicion = tmpParametroInforme.parametro;
	}
	//
	/**
	 * guarda los parametros de orden de compra (archivoParametrosOIT)
	 *
	 */
	public void guardarParametrosCarnet(){
		int numeroActualizaciones = 0;
		
		ParametroInforme tmpParametroInforme = new ParametroInforme(
																	"16",
																	anotacionesCarnetMaquinaInt,
																	miConexionDatos
																	);
		numeroActualizaciones += tmpParametroInforme.actualizar();
		
		//
		tmpParametroInforme = new ParametroInforme(
																	"17",
																	anotacionesCarnetPersonal,
																	miConexionDatos
																	);
		numeroActualizaciones += tmpParametroInforme.actualizar();
		
			
		//
		if(numeroActualizaciones == 2){
			JOptionPane.showMessageDialog(null,"Se han cambiado los parametros de carnet con exito");
		}else{
			JOptionPane.showMessageDialog(null,"Ha ocurrido un error en la modificacion del archivo de parametros de carnet");
		}
	}
	
	/**
	 * carga los parametros de orden de compra desde un archivo (archivoParametrosOIT)
	 *
	 */
	public void cargarParametrosCarnet(){
		ParametroInforme tmpParametroInforme = new ParametroInforme(
																null,
																miConexionDatos
																);
        
        tmpParametroInforme.cargar("codigo", "16");
        anotacionesCarnetMaquinaInt = tmpParametroInforme.parametro;
        
        //
        tmpParametroInforme.cargar("codigo", "17");
        anotacionesCarnetPersonal = tmpParametroInforme.parametro;
	}
	//
	/**
	 * guarda los parametros de costeo
	 *
	 */
	public void guardarParametrosCosteo(){
		int numeroActualizaciones = 0;
		
		ParametroInforme tmpParametroInforme = new ParametroInforme(
																	"30",
																	porcentajeCostosIndirectosFabricacion,
																	miConexionDatos
																	);
		numeroActualizaciones += tmpParametroInforme.actualizar();
		
		//
		tmpParametroInforme = new ParametroInforme(
																	"31",
																	porcentajeComisiones,
																	miConexionDatos
																	);
		numeroActualizaciones += tmpParametroInforme.actualizar();
		
		//
		tmpParametroInforme = new ParametroInforme(
																	"32",
																	porcentajeMargenUtilidad,
																	miConexionDatos
																	);
		numeroActualizaciones += tmpParametroInforme.actualizar();
			
		//
		if(numeroActualizaciones == 3){
			JOptionPane.showMessageDialog(null,"Se han cambiado los parametros de Costeo con exito");
		}else{
			JOptionPane.showMessageDialog(null,"Ha ocurrido un error en la modificacion del archivo de parametros de costeo");
		}
	}
	
	/**
	 * carga los parametros de costeo
	 *
	 */
	public void cargarParametrosCosteo(){
		ParametroInforme tmpParametroInforme = new ParametroInforme(
																null,
																miConexionDatos
																);
        
        tmpParametroInforme.cargar("codigo", "30");
        porcentajeCostosIndirectosFabricacion = tmpParametroInforme.parametro;
        
        //
        tmpParametroInforme.cargar("codigo", "31");
        porcentajeComisiones = tmpParametroInforme.parametro;
        
        //
        tmpParametroInforme.cargar("codigo", "32");
        porcentajeMargenUtilidad = tmpParametroInforme.parametro;
	}
	//
	/**
	 * carga el pie de pagina que se utilizara en los informes
	 *
	 */
	public void cargarPiePaginaInforme(){
		Empresa tmpEmpresa = new Empresa(
											null,
											miConexionDatos
											);
        
        tmpEmpresa.cargar("codigo", "1");
        
        piePaginaInforme = "Direccion: "+tmpEmpresa.direccion+" Tel: "+tmpEmpresa.telefono+" Fax: "+tmpEmpresa.fax+"\n"
        					+""+tmpEmpresa.sitioWeb+"   "+"e-mail: "+tmpEmpresa.email+"\n"
        					+tmpEmpresa.ciudad+" - "+tmpEmpresa.departamento+" - "+tmpEmpresa.pais;	
        
	}
	
	/**
	 * guarda los parametros de factura (archivoParametrosFactura)
	 *
	 */
	public void guardarParametrosFactura(){
		int numeroActualizaciones = 0;
		

		ParametroInforme tmpParametroInforme = new ParametroInforme(
													"20",
													textoLegalFactura,
													miConexionDatos
													);
		numeroActualizaciones += tmpParametroInforme.actualizar();
			
		//
		tmpParametroInforme = new ParametroInforme(
													"21",
													textoLegalNumeracionFactura,
													miConexionDatos
													);
		numeroActualizaciones += tmpParametroInforme.actualizar();
		
		//
		tmpParametroInforme = new ParametroInforme(
													"22",
													textoPiePaginaFactura,
													miConexionDatos
													);
		numeroActualizaciones += tmpParametroInforme.actualizar();
		
		//
		tmpParametroInforme = new ParametroInforme(
													"23",
													fuenteFactura,
													miConexionDatos
													);
		numeroActualizaciones += tmpParametroInforme.actualizar();
		
		//
		tmpParametroInforme = new ParametroInforme(
													"24",
													fuenteFacturaTextoLegal,
													miConexionDatos
													);
		numeroActualizaciones += tmpParametroInforme.actualizar();
		
		//
		tmpParametroInforme = new ParametroInforme(
													"25",
													fuenteFacturaTextoNumeracion,
													miConexionDatos
													);
		numeroActualizaciones += tmpParametroInforme.actualizar();
		
		//
		tmpParametroInforme = new ParametroInforme(
													"26",
													fuenteFacturaPiePagina,
													miConexionDatos
													);
		numeroActualizaciones += tmpParametroInforme.actualizar();
													
		//
		tmpParametroInforme = new ParametroInforme(
													"33",
													textoAdicionalFactura,
													miConexionDatos
													);
		numeroActualizaciones += tmpParametroInforme.actualizar();
													
		//
		tmpParametroInforme = new ParametroInforme(
													"34",
													fuenteFacturaTextoAdicional,
													miConexionDatos
													);
													
		numeroActualizaciones += tmpParametroInforme.actualizar();
		
		//
		if(numeroActualizaciones == 9){
			JOptionPane.showMessageDialog(null,"Se han cambiado los parametros de factura con exito");
		}else{
			JOptionPane.showMessageDialog(null,"Ha ocurrido un error en la modificacion del archivo de factura");
		}
	}
	
	/**
	 * carga los parametros de OrdenEjecucion desde un archivo (archivoParametrosOrdenEjecucion)
	 *
	 */
	public void cargarParametrosFactura(){
		ParametroInforme tmpParametroInforme = new ParametroInforme(
																null,
																miConexionDatos
																);
        
        //
        tmpParametroInforme.cargar("codigo", "20");
        textoLegalFactura = tmpParametroInforme.parametro;
        
        //
        tmpParametroInforme.cargar("codigo", "21");
        textoLegalNumeracionFactura = tmpParametroInforme.parametro;
        
        //
        tmpParametroInforme.cargar("codigo", "22");
        textoPiePaginaFactura = tmpParametroInforme.parametro;
        
        //
        tmpParametroInforme.cargar("codigo", "23");
        fuenteFactura = tmpParametroInforme.parametro;
        
        //
        tmpParametroInforme.cargar("codigo", "24");
        fuenteFacturaTextoLegal = tmpParametroInforme.parametro;
        
        //
        tmpParametroInforme.cargar("codigo", "25");
        fuenteFacturaTextoNumeracion = tmpParametroInforme.parametro;
        
        //
        tmpParametroInforme.cargar("codigo", "26");
        fuenteFacturaPiePagina = tmpParametroInforme.parametro;
        
        //
        tmpParametroInforme.cargar("codigo", "33");
        textoAdicionalFactura = tmpParametroInforme.parametro;
        
        //
        tmpParametroInforme.cargar("codigo", "34");
        fuenteFacturaTextoAdicional = tmpParametroInforme.parametro;
	}
	
	/**
	 * guarda los parametros de CC (archivoParametrosCC)
	 *
	 */
	public void guardarParametrosCC(){
		int numeroActualizaciones = 0;
		

		ParametroInforme tmpParametroInforme = new ParametroInforme(
													"40",
													textoLegalCC,
													miConexionDatos
													);
		numeroActualizaciones += tmpParametroInforme.actualizar();
			
		
		//
		tmpParametroInforme = new ParametroInforme(
													"41",
													textoPiePaginaCC,
													miConexionDatos
													);
		numeroActualizaciones += tmpParametroInforme.actualizar();
		
		//
		tmpParametroInforme = new ParametroInforme(
													"42",
													fuenteCC,
													miConexionDatos
													);
		numeroActualizaciones += tmpParametroInforme.actualizar();
		
		//
		tmpParametroInforme = new ParametroInforme(
													"43",
													fuenteCCTextoLegal,
													miConexionDatos
													);
		numeroActualizaciones += tmpParametroInforme.actualizar();
		
		//
		tmpParametroInforme = new ParametroInforme(
													"44",
													fuenteCCPiePagina,
													miConexionDatos
													);
		numeroActualizaciones += tmpParametroInforme.actualizar();
													
		//
		tmpParametroInforme = new ParametroInforme(
													"45",
													textoAdicionalCC,
													miConexionDatos
													);
		numeroActualizaciones += tmpParametroInforme.actualizar();
													
		//
		tmpParametroInforme = new ParametroInforme(
													"46",
													fuenteCCTextoAdicional,
													miConexionDatos
													);
													
		numeroActualizaciones += tmpParametroInforme.actualizar();
		
		//
		if(numeroActualizaciones == 7){
			JOptionPane.showMessageDialog(null,"Se han cambiado los parametros de cuenta de cobro con exito");
		}else{
			JOptionPane.showMessageDialog(null,"Ha ocurrido un error en la modificacion del archivo de cuenta de cobro");
		}
	}
	
	/**
	 * carga los parametros de OrdenEjecucion desde un archivo (archivoParametrosOrdenEjecucion)
	 *
	 */
	public void cargarParametrosCC(){
		ParametroInforme tmpParametroInforme = new ParametroInforme(
																null,
																miConexionDatos
																);
        
        //
        tmpParametroInforme.cargar("codigo", "40");
        textoLegalCC = tmpParametroInforme.parametro;
        
        //
        tmpParametroInforme.cargar("codigo", "41");
        textoPiePaginaCC = tmpParametroInforme.parametro;
        
        //
        tmpParametroInforme.cargar("codigo", "42");
        fuenteCC = tmpParametroInforme.parametro;
        
        //
        tmpParametroInforme.cargar("codigo", "43");
        fuenteCCTextoLegal = tmpParametroInforme.parametro;
        
        //
        tmpParametroInforme.cargar("codigo", "44");
        fuenteCCPiePagina = tmpParametroInforme.parametro;
        
        //
        tmpParametroInforme.cargar("codigo", "45");
        textoAdicionalCC = tmpParametroInforme.parametro;
        
        //
        tmpParametroInforme.cargar("codigo", "46");
        fuenteCCTextoAdicional = tmpParametroInforme.parametro;
	}
	/**
	 * guarda los parametros de numeracion de factura (archivoParametrosFactura)
	 *
	 */
	public void guardarParametrosNumeracionFactura(){
		int numeroActualizaciones = 0;
		
		ParametroInforme tmpParametroInforme = new ParametroInforme(
																	"27",
																	inicioNumeracion,
																	miConexionDatos
																	);
		numeroActualizaciones += tmpParametroInforme.actualizar();
		
		//
		tmpParametroInforme = new ParametroInforme(
													"28",
													finNumeracion,
													miConexionDatos
													);
		numeroActualizaciones += tmpParametroInforme.actualizar();
		
	
		
		//
		if(numeroActualizaciones == 2){
			JOptionPane.showMessageDialog(null,"Se han cambiado los parametros numeración de factura con exito");
		}else{
			JOptionPane.showMessageDialog(null,"Ha ocurrido un error en la modificacion del archivo de factura");
		}
	}
	
	/**
	 * carga los parametros de numeracion de factura desde un archivo (archivoParametrosOrdenEjecucion)
	 *
	 */
	public void cargarParametrosNumeracionFactura(){
		ParametroInforme tmpParametroInforme = new ParametroInforme(
																null,
																miConexionDatos
																);
        
        tmpParametroInforme.cargar("codigo", "27");
        inicioNumeracion = tmpParametroInforme.parametro;
        
        //
        tmpParametroInforme.cargar("codigo", "28");
        finNumeracion = tmpParametroInforme.parametro;

	}
	/////////////////metodos de carga password////////////////////
	/////////////////////////////////////////////////////////////////////////////////////////////////
	/**
	 * carga el password de configuracion desde un archivo (archivoPassword)
	 *
	 */
	public void cargarPasswordConfiguracion(){
		try{
    		ObjectInputStream in = new ObjectInputStream(new FileInputStream(archivoPassword));
            //carga los parametros
            passwordConfiguracion = (String)in.readObject();
            
            in.close();
       	} catch (ClassNotFoundException f) {
       	    System.out.println(f);
       	} catch (IOException f) {
       	    System.out.println(f);
       	    JOptionPane.showMessageDialog(null,"El archivo de configuracion no se encuentra");
       	}
	}
	/**
	 * guarda el password de configuracion en un archivo (archivoPassword)
	 *
	 */
	public void guardarPasswordConfiguracion(){
		try{
     		ObjectOutputStream out = new ObjectOutputStream(new FileOutputStream(archivoPassword));
     		
//     		passwordConfiguracion = "123456";	
			//escribe los parametros
			out.writeObject(passwordConfiguracion);
			
			//
    		out.close();
    	}catch (IOException f) {
        	System.out.println(f);
        	JOptionPane.showMessageDialog(null,"Ha ocurrido un error en la modificacion del archivo de password");
    	}
	}
	//
	/**
	 * guarda los parametros de registro de OIT
	 *
	 */
	public void guardarParametrosRegistroOIT(){
		int numeroActualizaciones = 0;
		
		ParametroInforme tmpParametroInforme = new ParametroInforme(
																	"18",
																	segundosRegistro,
																	miConexionDatos
																	);
		numeroActualizaciones += tmpParametroInforme.actualizar();
		
		//
		tmpParametroInforme = new ParametroInforme(
																	"19",
																	segundosPresentacionOIT,
																	miConexionDatos
																	);
		numeroActualizaciones += tmpParametroInforme.actualizar();
		
			
		//
		if(numeroActualizaciones == 2){
			JOptionPane.showMessageDialog(null,"Se han cambiado los parametros de registros de OIT con exito");
		}else{
			JOptionPane.showMessageDialog(null,"Ha ocurrido un error en la modificacion del archivo de parametros de registros de OIT");
		}
	}
	
	/**
	 * carga los parametros de registros de OIT
	 *
	 */
	public void cargarParametrosRegistroOIT(){
		ParametroInforme tmpParametroInforme = new ParametroInforme(
																null,
																miConexionDatos
																);
        
        tmpParametroInforme.cargar("codigo", "18");
        segundosRegistro = tmpParametroInforme.parametro;
        
        //
        tmpParametroInforme.cargar("codigo", "19");
        segundosPresentacionOIT = tmpParametroInforme.parametro;
	}
	//
	/**
	 * guarda los parametros de registro de OIT
	 *
	 */
	public void guardarParametrosArchivoFactura(){
		int numeroActualizaciones = 0;
		
		ParametroInforme tmpParametroInforme = new ParametroInforme(
																	"29",
																	prefijoArchivoFactura,
																	miConexionDatos
																	);
		numeroActualizaciones += tmpParametroInforme.actualizar();
		
			
		//
		if(numeroActualizaciones == 1){
			JOptionPane.showMessageDialog(null,"Se ha cambiado el prefijo de archivo de facturas con exito");
		}else{
			JOptionPane.showMessageDialog(null,"Ha ocurrido un error en la modificacion del prefijo de archivo de facturas");
		}
	}
	
	/**
	 * carga los parametros de registros de OIT
	 *
	 */
	public void cargarParametrosArchivoFactura(){
		ParametroInforme tmpParametroInforme = new ParametroInforme(
																null,
																miConexionDatos
																);
        
        tmpParametroInforme.cargar("codigo", "29");
        prefijoArchivoFactura = tmpParametroInforme.parametro;
	}
	
	/**
	 * guarda el parametro de seleccion de impresion de factura
	 *
	 */
	public void guardarParametrosImpresionFactura(){
		int numeroActualizaciones = 0;
		
		ParametroInforme tmpParametroInforme = new ParametroInforme(
																	"35",
																	idImpresionFactura,
																	miConexionDatos
																	);
		numeroActualizaciones += tmpParametroInforme.actualizar();
		
			
		//
		if(numeroActualizaciones == 1){
			JOptionPane.showMessageDialog(null,"Se ha cambiado la forma de imprimir la factura");
		}else{
			JOptionPane.showMessageDialog(null,"Ha ocurrido un error en la modificacion de la forma de impresion de factura");
		}
	}
	
	/**
	 * carga los parametros de registros de OIT
	 *
	 */
	public void cargarParametrosImpresionFactura(){
		ParametroInforme tmpParametroInforme = new ParametroInforme(
																null,
																miConexionDatos
																);
        
        tmpParametroInforme.cargar("codigo", "35");
        idImpresionFactura = tmpParametroInforme.parametro;
	}
	/**
	 * guarda el parametro de seleccion de impresion de factura
	 *
	 */
	public void guardarParametrosBloqueoCliente(){
		int numeroActualizaciones = 0;
		
		ParametroInforme tmpParametroInforme = new ParametroInforme(
																	"47",
																	numeroDiasBloqueo,
																	miConexionDatos
																	);
		numeroActualizaciones += tmpParametroInforme.actualizar();
		
			
		//
		if(numeroActualizaciones == 1){
			JOptionPane.showMessageDialog(null,"Se ha cambiado la forma de imprimir la factura");
		}else{
			JOptionPane.showMessageDialog(null,"Ha ocurrido un error en la modificacion de la forma de impresion de factura");
		}
	}
	
	/**
	 * carga los parametros de registros de OIT
	 *
	 */
	public void cargarParametrosBloqueoCliente(){
		ParametroInforme tmpParametroInforme = new ParametroInforme(
																null,
																miConexionDatos
																);
        
        tmpParametroInforme.cargar("codigo", "47");
        numeroDiasBloqueo = tmpParametroInforme.parametro;
	}
	/**
	 * permite eliminar todos los archivos en un directorio (usado para eliminar los archivos temporales de impresion)
	 *
	 */
	public static void borrarContenidoDirectorio(String inDirectorio){ 
		File miFile = new File(inDirectorio);
		File[] arrFiles = miFile.listFiles();
		
		for(int idx1 = 0; idx1 < arrFiles.length; idx1++){
			arrFiles[idx1].delete();
		}
		
	}
	
	/**
	 * permite eliminar todos los archivos en un directorio (usado para eliminar los archivos temporales de impresion)
	 *
	 */
	public static String generarNombreArchivoTemporal(){ 
		Calendar fecha = Calendar.getInstance();
		
		int fechaDia = fecha.get(Calendar.DAY_OF_MONTH);
		int fechaMes = fecha.get(Calendar.MONTH)+1;
		int fechaYear = fecha.get(Calendar.YEAR);
		
		SimpleDateFormat formateador = new SimpleDateFormat("kkmmssSSS");
		
		
		return "documento "+fechaYear+"-"+fechaMes+"-"+fechaDia+", "+formateador.format(fecha.getTime());
	}
	/**
	 * permite eliminar todos los archivos en un directorio (usado para eliminar los archivos temporales de impresion)
	 *
	 */
	public static String establecerNombreArchivo(ConexionDatos inConexionDatos){ 
		String outNombreArchivo = "";
		
		FileReaderWriter miFileReaderWriter = new FileReaderWriter(inConexionDatos);
		miFileReaderWriter.cargarParametrosInforme();
		
		if(miFileReaderWriter.opcionGuardarPDF.equals("1")){//impresion directa
			outNombreArchivo = FileReaderWriter.directorioTemporal+FileReaderWriter.generarNombreArchivoTemporal()+".pdf";
		}else{
			//Selecciona el nombre y ubicacion del archivo generado
			Frame frame = new Frame();
	    	FileDialog chooser = new FileDialog(frame,"Guardar informe DE");
	    	chooser.setMode(FileDialog.SAVE);
	    	chooser.setVisible(true);
	    	
	    	//lectura de la direccion del informe generado
			if(!chooser.getDirectory().equals("null")){
				outNombreArchivo = chooser.getDirectory() + chooser.getFile();
				if(!outNombreArchivo.endsWith(".pdf")){
					outNombreArchivo +=  ".pdf";
				}
			}	
		}
		
		return outNombreArchivo;
	}
	/**
	 * permite eliminar todos los archivos en un directorio (usado para eliminar los archivos temporales de impresion)
	 *
	 */
	public static String establecerNombreArchivo(ConexionDatos inConexionDatos, String inNombreArchivo){ 
		String outNombreArchivo = "";
		
		FileReaderWriter miFileReaderWriter = new FileReaderWriter(inConexionDatos);
		miFileReaderWriter.cargarParametrosInforme();
		
		if(miFileReaderWriter.opcionGuardarPDF.equals("1")){//impresion directa
			outNombreArchivo = FileReaderWriter.directorioTemporal+inNombreArchivo+".pdf";
		}else{
			//Selecciona el nombre y ubicacion del archivo generado
			Frame frame = new Frame();
	    	FileDialog chooser = new FileDialog(frame,"Guardar informe DE");
	    	chooser.setMode(FileDialog.SAVE);
	    	chooser.setVisible(true);
	    	
	    	//lectura de la direccion del informe generado
			if(!chooser.getDirectory().equals("null")){
				outNombreArchivo = chooser.getDirectory() + chooser.getFile();
				if(!outNombreArchivo.endsWith(".pdf")){
					outNombreArchivo +=  ".pdf";
				}
			}	
		}
		
		return outNombreArchivo;
	}
};