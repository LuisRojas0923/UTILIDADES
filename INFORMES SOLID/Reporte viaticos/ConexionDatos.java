/*
 Proyecto SOLID-ERP
 
 Clase 				ConexionDatos
 Archivo			ConexionDatos.java
  
 Fecha Creación 	1 de Octubre de 2019
 Autor  			Marlon Ramirez Jaime 1.130.637.509
 email 				marlonrj@gmail.com
*/

import java.awt.*;
import javax.swing.*;
import java.sql.*;
import java.io.*;
import java.util.*;

/**
 * Esta clase provee el acceso a diferentes manejadores de bases de datos y estable una conexion 
 * permanente con ellos para consulta y actualizacion
 *
 */
public class ConexionDatos {
		
	//atributos
	
	/** conexion */
	public Connection conexion;
	
	/** declaracion */
	public Statement stmt;
	
	/** almacena la direccion y puerto de conexion */
	public String direccionPuerto = ""; 
	
	/** almacena el nombre de la base de datos */	
	public String nombreBD = "";
	
	/** almacena el nombre de usuario */
	public String usuario = "";
	
	/** almacena el password */
	public String password = ""; 
		
	/** 
	 *constructor por defecto
	 */
	public ConexionDatos(){
		
	}
	/**
	 * metodo para conectarse a DBM Oracle
	 * --ej:		ConexionDatos.conexionOracle("127.0.0.1:1521","XE","user1","password1")
	 * @author Marlon Ramirez Jaime
	 * @param direccionPuerto -ip del equipo con la base de datos o localhost con el numero de puerto de conexion
	 * @param nombreBD -nombre de la base de datos requerida
	 * @param usuario -nombre de usuario para iniciar en el servidor 
	 * @param password -contraseña de usuario
	 *
	 * @return -devuelve 1 si la conexion fue exitosa, 0 si no lo fue
	 */
	//metodo para conectarse a Oracle 	(direccionPuerto : ip del equipo con la base de datos o localhost con el numero de puerto de conexion)
	//									(nombreBD : nombre de la base de datos requerida)
	//									(usuario : nombre de usuario para iniciar en el servidor)
	//									(password : contraseña de usuario)
	//
	public int conexionOracle(String inDireccionPuerto, String inNombreBD, String inUsuario, String inPassword){
			
		direccionPuerto = inDireccionPuerto;
		nombreBD = inNombreBD;
		usuario = inUsuario;
		password = inPassword;
			
		int comprobacion = 1;//con este flag se sabe si la operacion fue exitosa con 1 fue exitosa 0 no fue exitosa
		
		try{
  			String driver = "oracle.jdbc.driver.OracleDriver";
  			Class.forName( driver );
  			String url = "jdbc:oracle:thin:@"+direccionPuerto+":"+nombreBD; 
  			//ejemplos
  			//String url = "jdbc:oracle:thin:@127.0.0.1:1521:XE"; //localhost
			//String url = "jdbc:oracle:thin:@192.168.131.93:1521:ictus";
			//String url = "jdbc:oracle://ictus:puerto/basedatos";
			conexion = DriverManager.getConnection( url, usuario, password );				
			stmt = conexion.createStatement();
			//alteracion de la sesion para el formato de fecha "yyyy-mm-dd"
			stmt.executeQuery("alter session set nls_date_format = 'yyyy-mm-dd'");
		}catch( Exception x ) {
  			x.printStackTrace();
  			JOptionPane.showMessageDialog(null,"Error al conectarse a la base de datos Oracle en (CD)\n");
  			comprobacion = 0;
		}
		return comprobacion;// 1 exitosa 0 no fue exitosa
	}
	/**
	 * metodo para conectarse a DBM Access
	 *
	 * --ej:		ConexionDatos.conexionAccess("directorio", "password1")
	 *
	 * @param nombreOrigen -nombre del archivo .mdb ya registrado
	 * @param password -contraseña de usuario
	 *
	 * @return -devuelve 1 si la conexion fue exitosa, 0 si no lo fue
	 */
	//metodo para conectarse a access 	(nombreOrigen : nombre del archivo .mdb ya registrado)
	//									(password : contraseña de usuario)
	public int conexionAccess(String nombreOrigen ,String password){
			int comprobacion = 1;//con este flag se sabe si la operacion fue exitosa con 1 fue exitosa 0 no fue exitosa
			//conexion access
			try{
				String driver = "sun.jdbc.odbc.JdbcOdbcDriver";
      			Class.forName( driver );
      			String url = "jdbc:odbc:"+nombreOrigen; 
      			conexion = DriverManager.getConnection( url, password, password );				
				stmt = conexion.createStatement();
    		}catch( Exception x ) {
      				x.printStackTrace();
      				JOptionPane.showMessageDialog(null,"Error al conectarse a la base de datos access en (CD)\n");
      				comprobacion = 0;
    				}
			return comprobacion;// 1 exitosa 0 no fue exitosa
	}
	/**
	 * metodo para conectarse a DBM Mysql
	 * --ej:		conexionDatos.conexionMySQL("localhost:3306","baseDatosEjemplo","user1","password1")
	 *
	 * @param direccionPuerto -ip del equipo con la base de datos o localhost con el numero de puerto de conexion
	 * @param nombreBD -nombre de la base de datos requerida
	 * @param usuario -nombre de usuario para iniciar en el servidor 
	 * @param password -contraseña de usuario
	 *
	 * @return -devuelve 1 si la conexion fue exitosa, 0 si no lo fue
	 */
	//metodo para conectarse a MySQL 	(direccionPuerto : ip del equipo con la base de datos o localhost con el numero de puerto de conexion)
	//									(nombreBD : nombre de la base de datos requerida)
	//									(usuario : nombre de usuario para iniciar en el servidor)
	//									(password : contraseña de usuario)
	//ej:		conexionDatos.conexionMySQL("localhost:3306","baseDatosEjemplo","user1","password1")
	public int conexionMySQL(String inDireccionPuerto, String inNombreBD, String inUsuario, String inPassword){
		
		direccionPuerto = inDireccionPuerto;
		nombreBD = inNombreBD;
		usuario = inUsuario;
		password = inPassword;
		
		int comprobacion = 1;//con este flag se sabe si la operacion fue exitosa con 1 fue exitosa 0 no fue exitosa
		
		try{
  			String driver = "com.mysql.jdbc.Driver";
  			Class.forName( driver );
  			String url = "jdbc:mysql://"+direccionPuerto+"/"+nombreBD; 
  			//String url = "jdbc:mysql://"+"192.168.3.3:3306"+"/"+nombreBD; 
  			//ejemplos
  			//String url = "jdbc:mysql://localhost:3306/test"; 
			//String url = "jdbc:mysql://malpelo.eisc.univalle.edu.co/marlonrj"; 
			conexion = DriverManager.getConnection( url, usuario, password );				
			stmt = conexion.createStatement();
		}catch( Exception x ) {
			x.printStackTrace();
			JOptionPane.showMessageDialog(null,"Error al conectarse a la base de datos MySQL en (CD)\n");
			comprobacion = 0;
		}
		return comprobacion;// 1 exitosa 0 no fue exitosa
	}
	/**
	 * metodo para conectarse a DBM PostgreSQL
	 * --ej:		conexionDatos.conexionPostgreSQL("localhost:3306","baseDatosEjemplo","user1","password1")
	 *
	 * @param direccionPuerto -ip del equipo con la base de datos o localhost con el numero de puerto de conexion
	 * @param nombreBD -nombre de la base de datos requerida
	 * @param usuario -nombre de usuario para iniciar en el servidor 
	 * @param password -contraseña de usuario
	 *
	 * @return -devuelve 1 si la conexion fue exitosa, 0 si no lo fue
	 */
	//metodo para conectarse a PostgreSQL 	(direccionPuerto : ip del equipo con la base de datos o localhost con el numero de puerto de conexion)
	//									(nombreBD : nombre de la base de datos requerida)
	//									(usuario : nombre de usuario para iniciar en el servidor)
	//									(password : contraseña base de datos access)
	//ej:		conexionDatos.conexionMySQL("localhost:3306","baseDatosEjemplo","user1","password1")
	public int conexionPostgreSQL(String inDireccionPuerto, String inNombreBD, String inUsuario, String inPassword){
		
		direccionPuerto = inDireccionPuerto;
		nombreBD = inNombreBD;
		usuario = inUsuario;
		password = inPassword;
		
		int comprobacion = 1;//con este flag se sabe si la operacion fue exitosa con 1 fue exitosa 0 no fue exitosa
		
		try{
  			String driver = "org.postgresql.Driver";
  			Class.forName( driver );
  			String url = "jdbc:postgresql://"+direccionPuerto+"/"+nombreBD; 
  			//String url = "jdbc:mysql://"+"192.168.3.3:3306"+"/"+nombreBD; 
  			//ejemplos
  			//String url = "jdbc:mysql://localhost:3306/test"; 
			//String url = "jdbc:mysql://malpelo.eisc.univalle.edu.co/marlonrj"; 
			conexion = DriverManager.getConnection( url, usuario, password );				
			stmt = conexion.createStatement(ResultSet.TYPE_SCROLL_INSENSITIVE,
                                  ResultSet.CONCUR_UPDATABLE);
		}catch( Exception x ) {
			x.printStackTrace();
			JOptionPane.showMessageDialog(null,"Error al conectarse a la base de datos postgreSQL en (CD)\n");
			comprobacion = 0;
		}
		return comprobacion;// 1 exitosa 0 no fue exitosa
	}
	/**
	 * metodo para conectarse a DBM PostgreSQL no muestra mensajes de confirmacion o falla
	 * --ej:		conexionDatos.conexionPostgreSQL("localhost:3306","baseDatosEjemplo","user1","password1")
	 *
	 * @param direccionPuerto -ip del equipo con la base de datos o localhost con el numero de puerto de conexion
	 * @param nombreBD -nombre de la base de datos requerida
	 * @param usuario -nombre de usuario para iniciar en el servidor 
	 * @param password -contraseña de usuario
	 *
	 * @return -devuelve 1 si la conexion fue exitosa, 0 si no lo fue
	 */
	//metodo para conectarse a PostgreSQL 	(direccionPuerto : ip del equipo con la base de datos o localhost con el numero de puerto de conexion)
	//									(nombreBD : nombre de la base de datos requerida)
	//									(usuario : nombre de usuario para iniciar en el servidor)
	//									(password : contraseña base de datos access)
	//ej:		conexionDatos.conexionMySQL("localhost:3306","baseDatosEjemplo","user1","password1")
	public int conexionPostgreSQL2(String direccionPuerto, String nombreBD, String usuario, String password){
			int comprobacion = 1;//con este flag se sabe si la operacion fue exitosa con 1 fue exitosa 0 no fue exitosa
			//conexion MySQL
			try{
      			String driver = "org.postgresql.Driver";
	  			Class.forName( driver );
	  			String url = "jdbc:postgresql://"+direccionPuerto+"/"+nombreBD; 
	  			//String url = "jdbc:mysql://"+"192.168.3.3:3306"+"/"+nombreBD; 
	  			//ejemplos
	  			//String url = "jdbc:mysql://localhost:3306/test"; 
				//String url = "jdbc:mysql://malpelo.eisc.univalle.edu.co/marlonrj"; 
    			conexion = DriverManager.getConnection( url, usuario, password );				
				stmt = conexion.createStatement(ResultSet.TYPE_SCROLL_INSENSITIVE,
                                      ResultSet.CONCUR_UPDATABLE);
    		}catch( Exception x ) {
      				x.printStackTrace();
//      				JOptionPane.showMessageDialog(null,"Error al conectarse a la base de datos postgreSQL en (CD)\n");
      				comprobacion = 0;
    		}
    		return comprobacion;// 1 exitosa 0 no fue exitosa
	}
	/**
	 * metodo que ejecuta una sentencia SQL de consulta, muestra mensaje de error en interfaz
	 * --ej:		conexionDatos.funcionConsultar("select * from tabla")
	 *
	 * @param comandoSql -sentencia SQL a ejecutar
	 * 
	 * @return -devuelve 1 si la operacion fue exitosa, 0 si no lo fue
	 */
	//funcion consultar
	public ResultSet funcionConsultar(String comandoSql){
		ResultSet resultSet = null;
		System.out.println(comandoSql);
		try{
			
			//realizando consulta
			resultSet = stmt.executeQuery(comandoSql);
			
			return resultSet;
			
		}catch( Exception x ) {
    	 	x.printStackTrace();
     		JOptionPane.showMessageDialog(null,"No se puede realizar la consulta deseada en (CD)\n"+x.getMessage());
     		return resultSet;
   		}
   		
    }
    /**
	 * metodo que ejecuta una sentencia SQL de consulta, no muestra mensaje de excepcion en GUI
	 * --ej:		conexionDatos.funcionConsultar("select * from tabla")
	 *
	 * @param comandoSql -sentencia SQL a ejecutar
	 * 
	 * @return -devuelve 1 si la operacion fue exitosa, 0 si no lo fue
	 */
	//funcion consultar
	public ResultSet funcionConsultar2(String comandoSql){
		ResultSet resultSet = null;
		System.out.println(comandoSql);
		try{
			
			//realizando consulta
			resultSet = stmt.executeQuery(comandoSql);
			
			return resultSet;
			
		}catch( Exception x ) {
//////    	 	x.printStackTrace();
//////     		JOptionPane.showMessageDialog(null,"No se puede realizar la consulta deseada en (CD)\n"+x.getMessage());
     		return resultSet;
   		}
   		
    }
    /**
	 * metodo que ejecuta una sentencia SQL de modificacion, muestra mensaje de error en interfaz 
	 * --ej:		conexionDatos.funcionActualizar("drop table tabla1")
	 *
	 * @param comandoSql -sentencia SQL a ejecutar
	 * 
	 * @return -devuelve 1 si la operacion fue exitosa, 0 si no lo fue
	 */
    //funcion modificar/actualizar y crear
    public int funcionActualizar(String comandoSql){
    	int comprobacion = 1;//con este flag se sabe si la operacion fue exitosa con 1 fue exitosa 0 no fue exitosa
 		System.out.println(comandoSql);
    	try{
		
			stmt.executeUpdate(comandoSql); // ejecuta el SQL
			
		}catch( Exception x ) {
    	 	x.printStackTrace();
    		JOptionPane.showMessageDialog(null,"Introduccion de datos fallida en (CD)\n"+x.getMessage());
    		comprobacion = 0;
    	}
		if(comprobacion==1){
			//en esta version no muestra mensaje de confirmacion
			//JOptionPane.showMessageDialog(null,"Introduccion de datos exitosa");
		}
    	
    	return comprobacion; // 1 exitosa 0 no fue exitosa
    }
    /**
	 * metodo que ejecuta una sentencia SQL de modificacion, no muestra mensaje de excepcion en GUI
	 * --ej:		conexionDatos.funcionActualizar2("drop table tabla1")
	 *
	 * @param comandoSql -sentencia SQL a ejecutar
	 * 
	 * @return -devuelve 1 si la operacion fue exitosa, 0 si no lo fue
	 */
    //funcion modificar/actualizar y crear es igual a la anterior solo que no muestra excepciones en mensaje
    public int funcionActualizar2(String comandoSql){
    	int comprobacion = 1;//con este flag se sabe si la operacion fue exitosa con 1 fue exitosa 0 no fue exitosa
 		System.out.println(comandoSql);
    	try{
			
			stmt.executeUpdate(comandoSql); // ejecuta el SQL
	
		}catch( Exception x ) {
    	 		x.printStackTrace();
    			comprobacion = 0;
    	}
		if(comprobacion==1){
			//en esta version no muestra mensaje de confirmacion
			//JOptionPane.showMessageDialog(null,"Introduccion de datos exitosa");
		}
    	
    	return comprobacion; // 1 exitosa 0 no fue exitosa
    }
    /**
	 * metodo que finaliza la conexion
	 * --ej:		conexionDatos.terminarConexion()
	 * 
	 * @return -devuelve 1 si la operacion fue exitosa, 0 si no lo fue
	 */
    //funcion que finaliza la conexion
    public int terminarConexion(){
    	int comprobacion = 1;//con este flag se sabe si la operacion fue exitosa con 1 fue exitosa 0 no fue exitosa
    	try{
    		conexion.close(); 
    	}catch(Exception x){
    		x.printStackTrace();
    		comprobacion = 0;
    	}
    	return comprobacion; // 1 exitosa 0 no fue exitosa	
    }
    /**
	 * metodo que arma una sentencia SQL con los paraemtros de entrada
	 * --ej:
	 *
	 * @param inNombresTablas -nombres de las tablas a consultar
	 * @param inNombresCampos -nombres de los campos a consultar
	 * @param inAlias -alias que se asignaran
	 * @param inCamposJoin1 -campos que haran join
	 * @param inCamposJoin2 -campos que haran join
	 * @param inOperador -Operadores usados para join
	 * @param inOperadorLogico -Operadores logicos usados para unir verificaciones y sentencias en el where
	 *
	 * @return -retorna el ResultSet de la consulta
	 */
	public String armarSentencia(			String[] inNombresTablas,
											String[] inNombreCampos,
											String[] inAlias,
											String[] inCamposJoin1,
											String[] inCamposJoin2,
											String[] inOperador,
											String[] inOperadorLogico
													){
														
		String consulta = "";
		
		//armado de la consulta
		
		//segmento que arma la consulta en el select
//		consulta = "select distinct ";
		consulta = "select ";
				
		for(int idx1 = 0; idx1 < inNombreCampos.length; idx1++ ){
			if(idx1==inNombreCampos.length-1){//si es el ultimo campo
				consulta = consulta + inNombreCampos[idx1] +" as \""+inAlias[idx1]+"\" ";	
			}else{
				consulta = consulta + inNombreCampos[idx1] +" as \""+inAlias[idx1]+"\", ";
			}
		}
		
		//segmento que arma la consulta en el from
		consulta = consulta + " from ";
		
		for(int idx1 = 0; idx1 < inNombresTablas.length; idx1++ ){
			if(idx1==inNombresTablas.length-1){//si es el ultimo nombre de tabla
				consulta = consulta + inNombresTablas[idx1]+" ";	
			}else{
				consulta = consulta + inNombresTablas[idx1]+", ";
			}
		}
		
		//segmento que arma la consulta en el where
		if(inCamposJoin1!=null){//verifica si la consulta no tiene restricciones 
		
			consulta = consulta + " where ";
			
			for(int idx1 = 0; idx1 < inCamposJoin1.length; idx1++ ){
				if(idx1==inCamposJoin1.length-1){//si es el ultimo campo de join
					consulta = consulta + inCamposJoin1[idx1]+" "+inOperador[idx1]+" "+inCamposJoin2[idx1]+" ";	
				}else{
					consulta = consulta + inCamposJoin1[idx1]+" "+inOperador[idx1]+" "+inCamposJoin2[idx1]+" "+inOperadorLogico[idx1]+" ";	
				}
			}
		}
		
		//cualquier cambio a consulta se debe hacer aqui
		consulta = consulta + "";
		
		System.out.println(consulta);												
														
		return consulta;
	}
    /**
	 * metodo que ejecuta una sentencia que arma a partir de datos pasados en arreglos de cadenas, muestra mensaje de error en interfaz
	 * --ej:
	 *
	 * @param inNombresTablas -nombres de las tablas a consultar
	 * @param inNombresCampos -nombres de los campos a consultar
	 * @param inAlias -alias que se asignaran
	 * @param inCamposJoin1 -campos que haran join
	 * @param inCamposJoin2 -campos que haran join
	 * @param inOperador -Operadores usados para join
	 * @param inOperadorLogico -Operadores logicos usados para unir verificaciones y sentencias en el where
	 *
	 * @return -retorna el ResultSet de la consulta
	 */
	public ResultSet funcionConsultarParametrizada(		String[] inNombresTablas,
														String[] inNombreCampos,
														String[] inAlias,
														String[] inCamposJoin1,
														String[] inCamposJoin2,
														String[] inOperador,
														String[] inOperadorLogico
													){
		ResultSet resultSet = null;
		String consulta = armarSentencia(	inNombresTablas,
											inNombreCampos,
											inAlias,
											inCamposJoin1,
											inCamposJoin2,
											inOperador,
											inOperadorLogico);
		
		
		
		try{
			
			//realizando consulta
			resultSet = stmt.executeQuery(consulta);
			
   			return resultSet;
			
			}catch( Exception x ) {
    	 		x.printStackTrace();
     			JOptionPane.showMessageDialog(null,"No se puede realizar la consulta deseada en (CD)\n"+x.getMessage());
     			return resultSet;
   			}	
    }
    /**
	 * metodo que ejecuta una sentencia que arma a partir de datos pasados en arreglos de cadenas, muestra mensaje de error en interfaz
	 * --ej:
	 *
	 * @param inNombresTablas -nombres de las tablas a consultar
	 * @param inNombresCampos -nombres de los campos a consultar
	 * @param inAlias -alias que se asignaran
	 * @param inCamposJoin1 -campos que haran join
	 * @param inCamposJoin2 -campos que haran join
	 * @param inOperador -Operadores usados para join
	 * @param inOperadorLogico -Operadores logicos usados para unir verificaciones y sentencias en el where
	 * @param inNumeroFilasLimite -Numero de Filas Limitado
	 *
	 * @return -retorna el ResultSet de la consulta
	 */
	public ResultSet funcionConsultarParametrizadaLimitada(		String[] inNombresTablas,
														String[] inNombreCampos,
														String[] inAlias,
														String[] inCamposJoin1,
														String[] inCamposJoin2,
														String[] inOperador,
														String[] inOperadorLogico,
														String inNumeroFilasLimite
													){
		ResultSet resultSet = null;
		String consulta = armarSentencia(	inNombresTablas,
											inNombreCampos,
											inAlias,
											inCamposJoin1,
											inCamposJoin2,
											inOperador,
											inOperadorLogico);
		
		consulta = consulta + " limit "+inNumeroFilasLimite;
		
		System.out.println("CCCCCCCCCCCCCCCCCCCCCCCC "+consulta);
		
		try{
			
			//realizando consulta
			resultSet = stmt.executeQuery(consulta);
			
   			return resultSet;
			
			}catch( Exception x ) {
    	 		x.printStackTrace();
     			JOptionPane.showMessageDialog(null,"No se puede realizar la consulta deseada en (CD)\n"+x.getMessage());
     			return resultSet;
   			}	
    }
//////////    /**
//////////	 * crea una tabla y la coloca en un panel
//////////	 * --ej:		conexionDatos.funcionConsultar("select * from tabla")
//////////	 *
//////////	 * @param inPanel -panel en el que se coloca la tabla generada
//////////	 * @param inResulSet -resultado de la consulta realizada
//////////	 * @param inEncabezados -encabezados que tendra la tabla
//////////	 * @param inOpcionSeq -permite decidir si la tabla se mostrara con una secuencia determinada en la primera columna
//////////	 * 
//////////	 * @return -retorna la MTabla construida
//////////	 */
//////////	 public MTabla armarTabla( 	JPanel inPanel, 
//////////	 							ResultSet inResulSet, 
//////////	 							String[] inEncabezados, 
//////////	 							boolean inOpcionSeq, 
//////////	 							boolean inOpcionSel){
//////////		
//////////		
//////////		inPanel.setLayout(new GridLayout(1,1));
//////////		inPanel.removeAll();
//////////		
//////////		//alargando los encabezados dependiendo de la opcion
//////////		String[] encabezados;
//////////		
//////////		if(inOpcionSeq){
//////////			encabezados = new String[inEncabezados.length+1];
//////////			encabezados[0] = " ";
//////////			for(int idx = 0; idx < inEncabezados.length; idx++ ){
//////////				encabezados[idx + 1] = inEncabezados[idx];
//////////			}
//////////		}else{
//////////			encabezados = inEncabezados;
//////////		}
//////////		
//////////		
//////////		int longitud = 0;
//////////		int numeroColumnas = encabezados.length;
//////////		try{
//////////			if(inResulSet!=null){
//////////				while(inResulSet.next()){
//////////					longitud++;
//////////				}
//////////				inResulSet.beforeFirst();
//////////			}
//////////		}catch(Exception e){
//////////			e.printStackTrace();
//////////		}
//////////		MTabla tablaResultado = new MTabla(inPanel, longitud, numeroColumnas, 120, encabezados);
//////////		
//////////		if(inOpcionSeq){
//////////			tablaResultado.miTabla.getColumnModel().getColumn(0).setResizable(false);
//////////		}
//////////		
//////////		System.out.println(longitud);
//////////		//llenado de la tabla
//////////		
//////////		int anchos[] = new int[numeroColumnas];
//////////		 
//////////		
//////////		int pixelesLetra = 9;
//////////		int pixelesLetraSecuencia = 22;
//////////		int alturaCelda = 18;
//////////		
//////////		//coloca los anchos predeterminados
//////////		for(int idx = 0; idx < numeroColumnas; idx++ ){
//////////			anchos[idx] =  encabezados[idx].length()  * pixelesLetra;
//////////		}
//////////		
//////////		//
//////////		if(inOpcionSeq){
//////////			anchos[0] = 22;
//////////		}else{
//////////			anchos[0] = 54;
//////////		}
//////////		
//////////		//anchos[0] = 22;
//////////		
//////////		int idx1 = 0;
//////////		try{
//////////			if(inResulSet!=null){
//////////				while(inResulSet.next()){//filas
//////////					
//////////					if(inOpcionSeq){
//////////						tablaResultado.asignaValor(""+(idx1+1), idx1, 0);
//////////						for(int idx2 = 1; idx2 < numeroColumnas; idx2++ ){
//////////							String tmpPalabra = inResulSet.getString(idx2);
//////////							
//////////							tablaResultado.asignaValor(tmpPalabra,idx1,idx2);
//////////							
//////////							if((tmpPalabra.length()*pixelesLetra) > anchos[idx2]){
//////////								anchos[idx2] = (tmpPalabra.length()*pixelesLetra);
//////////							}
//////////						}
//////////						if(((""+(idx1+1)).length()*pixelesLetraSecuencia) > anchos[0]){
//////////							anchos[0] = ((""+(idx1+1)).length()*pixelesLetraSecuencia);
//////////						}
//////////					}else{
//////////						//tablaResultado.asignaValor(""+(idx1+1), idx1, 0);
//////////						for(int idx2 = 0; idx2 < numeroColumnas; idx2++ ){
//////////							String tmpPalabra = inResulSet.getString(idx2+1);
//////////							
//////////							tablaResultado.asignaValor(tmpPalabra,idx1,idx2);
//////////							
//////////							if((tmpPalabra.length()*pixelesLetra) > anchos[idx2]){
//////////								anchos[idx2] = (tmpPalabra.length()*pixelesLetra);
//////////							}
//////////						}
//////////					}
//////////					idx1++;
//////////				}
//////////				inResulSet.beforeFirst();
//////////			}
//////////		}catch(Exception e){
//////////			e.printStackTrace();
//////////		}
//////////			
//////////		//cambia formato de tabla
//////////		for(int idx3 = 0; idx3 < numeroColumnas; idx3++ ){
//////////			tablaResultado.setAnchoCelda(idx3,anchos[idx3]);
//////////		}
//////////		tablaResultado.asignarAlturaCelda(alturaCelda);
//////////		
//////////		
//////////		//cambio de formato de las tablas segun la opcion seleccionada
//////////		
//////////		if(!inOpcionSel){
//////////			JTextField miField = new JTextField();
//////////			miField.setEditable(false);
//////////			miField.setFocusable(false);
//////////			//miField2.setOpaque(false);
//////////			for(int idx4 = 0; idx4 < numeroColumnas; idx4++){
//////////				tablaResultado.miTabla.getColumnModel().getColumn(idx4).setCellEditor(new DefaultCellEditor(miField));
//////////			}
//////////		}
//////////		
//////////		//actualizacion de la interfaz	
//////////		inPanel.updateUI();
//////////		
//////////		return tablaResultado;
//////////
//////////	}
	/**
	 * crea una tabla y la coloca en un panel
	 * --ej:		conexionDatos.funcionConsultar("select * from tabla")
	 *
	 * @param inPanel -panel en el que se coloca la tabla generada
	 * @param inResulSet -resultado de la consulta realizada
	 * @param inEncabezados -encabezados que tendra la tabla
	 * @param inOpcionSeq -permite decidir si la tabla se mostrara con una secuencia determinada en la primera columna
	 * 
	 * @return -retorna la MTabla construida
	 */
	 public MTabla armarTabla( 	JPanel inPanel, 
	 							ResultSet inResulSet, 
	 							String[] inEncabezados, 
	 							boolean inOpcionSeq, 
	 							boolean inOpcionSel){
		
		
		inPanel.removeAll();
		inPanel.setLayout(new GridLayout(1,1));
		
		
		//alargando los encabezados dependiendo de la opcion (coloca el arreglo "inEncabezados" en el arreglo "encabezados")
		String[] encabezados;
		
		if(inOpcionSeq){
			encabezados = new String[inEncabezados.length+1];
			encabezados[0] = "  ";
			for(int idx = 0; idx < inEncabezados.length; idx++ ){
				encabezados[idx + 1] = inEncabezados[idx];
			}
		}else{
			encabezados = inEncabezados;
		}
		
		int longitud = 0;
		int numeroColumnas = encabezados.length;
		try{
			if(inResulSet!=null){
				while(inResulSet.next()){
					longitud++; //cuenta el numero de filas de la consulta
				}
				inResulSet.beforeFirst();
			}
		}catch(Exception e){
			e.printStackTrace();
		}
		
		//creacion de la tabla de resultados
		MTabla tablaResultado = new MTabla(inPanel, longitud, numeroColumnas, 120, encabezados);
		
		//si esta habilitada la opcion de secuencia coloca la columna de secuencia Resizable false
		if(inOpcionSeq){
			tablaResultado.miTabla.getColumnModel().getColumn(0).setResizable(false);
		}
		
		System.out.println(longitud);
		//llenado de la tabla
		
		String cadenasMaximaLongitud[] = new String[numeroColumnas];
		
		//
		int alturaCelda = 21;
		int longitudMaxima = 500;
		int longitudIteracion = 0;
		
		//coloca los anchos predeterminados
		for(int idx = 0; idx < numeroColumnas; idx++ ){

			cadenasMaximaLongitud[idx] = encabezados[idx];
		}
		
		int idx1 = 0;
		try{
			if(inResulSet!=null){
				while(inResulSet.next()){//filas
					
					if(inOpcionSeq){ //si tiene columna de numeracion
						tablaResultado.asignaValor(""+(idx1+1), idx1, 0);
						for(int idx2 = 1; idx2 < numeroColumnas; idx2++ ){ //comienza a leer el resultado
							String tmpPalabra = inResulSet.getString(idx2); //saca la palabra
							
							tablaResultado.asignaValor(tmpPalabra,idx1,idx2); //pone la palabra en la tabla
							
							//
							longitudIteracion = tmpPalabra.length();
							
							System.out.println("IMPRESION "+tmpPalabra+" "+tmpPalabra.length() );
							
							if( (longitudIteracion <= longitudMaxima) &&
								(longitudIteracion > cadenasMaximaLongitud[idx2].length())){ //si la palabra actual es mas larga que la mas larga anterior

								cadenasMaximaLongitud[idx2] = tmpPalabra;
							}
						}
						if((""+(idx1+1)).length() > cadenasMaximaLongitud[0].length()){

							cadenasMaximaLongitud[0] = ""+(idx1+1);
						} //hasta aqui voy
					}else{
						
						for(int idx2 = 0; idx2 < numeroColumnas; idx2++ ){ //comienza a leer el resultado
							String tmpPalabra = inResulSet.getString(idx2+1); //saca la palabra
							
							tablaResultado.asignaValor(tmpPalabra,idx1,idx2);
							
							longitudIteracion = tmpPalabra.length();
							if(	(longitudIteracion <= longitudMaxima) &&
								(longitudIteracion > cadenasMaximaLongitud[idx2].length())){

								cadenasMaximaLongitud[idx2] = tmpPalabra;
							}
						}
					}
					idx1++;
				}
				inResulSet.beforeFirst();
			}
		}catch(Exception e){
			e.printStackTrace();
		}
		 
		

			
		//cambia formato de tabla
		for(int idx3 = 0; idx3 < numeroColumnas; idx3++ ){
			tablaResultado.setAnchoCelda(idx3,Restriccion.calcularLongitudCadena(cadenasMaximaLongitud[idx3]));
		}
		tablaResultado.asignarAlturaCelda(alturaCelda);
		
		
		//cambio de formato de las tablas segun la opcion seleccionada
		
		if(!inOpcionSel){
			JTextField miField = new JTextField();
			miField.setEditable(false);
			miField.setFocusable(false);
			//miField2.setOpaque(false);
			for(int idx4 = 0; idx4 < numeroColumnas; idx4++){
				tablaResultado.miTabla.getColumnModel().getColumn(idx4).setCellEditor(new DefaultCellEditor(miField));
			}
		}
		
		//actualizacion de la interfaz	
		inPanel.updateUI();
		
		//
		System.out.println("Impresion de la cadena de maxima longitud");
		Restriccion.imprimirArreglo(cadenasMaximaLongitud);
		return tablaResultado;

	}
	/**
	 * crea un arreglo de String con los datos del resulset que es pasado como parametro
	 *
	 * @param inResultSet -resultset con los datos
	 * 
	 * @return -retorna el arreglo de String construido
	 *
	 */
	public static String[][] armarArreglo(ResultSet inResultSet){
		int nFilas = 0;
		int nColumnas = 0;
		
		//calcula el numero de filas
		try{
			nColumnas = inResultSet.getMetaData().getColumnCount();
			while(inResultSet.next()){
				nFilas++;
			}
			inResultSet.beforeFirst();
		}catch(Exception e){
			e.printStackTrace();	
		}
		
		//crea el arreglo con la longitud adecuada
		String[][] outArreglo = new String[nFilas][nColumnas];
		
		//llena el arreglo de salida
		try{
			int idx1 = 0;
			while(inResultSet.next()){
				for(int idx2 = 0; idx2 < nColumnas; idx2++){
					outArreglo[idx1][idx2] = inResultSet.getString(idx2+1);
				}
				idx1++;
			}
			inResultSet.beforeFirst();
		}catch(Exception e){
			e.printStackTrace();	
		}
		
		//devuelve el arreglo con los datos
		return outArreglo;
	}
	
	/**
	 * Obtiene el serial de la maquina que esta corriendo el programa
	 *
	 */
	public static String obtenerSerial(){
		
		//serial
		String serial = "";
		try{
		
			Process p = Runtime.getRuntime().exec("wmic bios get serialnumber"); 
			BufferedReader input = new BufferedReader(new InputStreamReader(p.getInputStream()));
   
    	  	//en windows 10 se leen las 3 lineas para obtener el serial   
    	  	serial = input.readLine();
    	  	serial = input.readLine();
    	  	serial = input.readLine();
    		  
  		    input.close();		
			
		}catch(Exception e){
			e.printStackTrace();
		}
		//devuelve serial sin blancos
		return serial.trim();
	}
	
	public static void safe(){
		//validacion serial de la maquina y tiempo de ejecucion
		Calendar calendario = Calendar.getInstance();
		String year = ""+calendario.get(Calendar.YEAR);
		String mes = ""+(calendario.get(Calendar.MONTH)+1);
		
		//pone los ceros al frente
		if(Integer.parseInt(mes)<10){
			mes = "0"+mes;
		}
		
//////		if(!ConexionDatos.obtenerSerial().equals("H4N0CX024175142j")){
//////			System.exit(0);
//////		}else 
			
		if(Integer.parseInt(""+year+mes) >202401){ //se ingresa el year luego mes con ceros al frente
			System.exit(0);
		}
		//termina serial
	}	
	
	
	
//	public MTabla armarTabla(JPanel inPanel, ResultSet inResulSet, String[] inEncabezados, boolean inOpcionSeq, boolean inOpcionSel){
//		
//		
//		inPanel.setLayout(new GridLayout(1,1));
//		inPanel.removeAll();
//		
//		//alargando los encabezados dependiendo de la opcion
//		String[] encabezados;
//		
//		if(inOpcionSeq){
//			encabezados = new String[inEncabezados.length+1];
//			encabezados[0] = " ";
//			for(int idx = 0; idx < inEncabezados.length; idx++ ){
//				encabezados[idx + 1] = inEncabezados[idx];
//			}
//		}else{
//			encabezados = inEncabezados;
//		}
//		
//		
//		int longitud = 0;
//		int numeroColumnas = encabezados.length;
//		try{
//			while(inResulSet.next()){
//			longitud++;
//			}
//			inResulSet.beforeFirst();
//		}catch(Exception e){
//			e.printStackTrace();
//			}
//		MTabla tablaResultado = new MTabla(inPanel, longitud, numeroColumnas, 120, encabezados);
//		
//		if(inOpcionSeq){
//			tablaResultado.miTabla.getColumnModel().getColumn(0).setResizable(false);
//		}
//		
//		System.out.println(longitud);
//		//llenado de la tabla
//		
//		int anchos[] = new int[numeroColumnas];
//		 
//		
//		int pixelesLetra = 9;
//		int pixelesLetraSecuencia = 22;
//		int alturaCelda = 18;
//		
//		//coloca los anchos predeterminados
//		for(int idx = 0; idx < numeroColumnas; idx++ ){
//			anchos[idx] =  Restriccion.getLengthString(encabezados[idx]);
//		}
//		
//		//
//		if(inOpcionSeq){
//			anchos[0] = 22;
//		}
//		
//		//anchos[0] = 22;
//		
//		int idx1 = 0;
//		try{
//			while(inResulSet.next()){//filas
//				
//				if(inOpcionSeq){
//					tablaResultado.asignaValor(""+(idx1+1), idx1, 0);
//					for(int idx2 = 1; idx2 < numeroColumnas; idx2++ ){
//						String tmpPalabra = inResulSet.getString(idx2);
//						
//						tablaResultado.asignaValor(tmpPalabra,idx1,idx2);
//						
//						//capturando los anchos de las columnas
//						int tmpLongitud = Restriccion.getLengthString(tmpPalabra);
//						if(tmpLongitud > anchos[idx2]){
//							anchos[idx2] = tmpLongitud;
//						}
//					}
//					int tmpLongitud = Restriccion.getLengthString(""+(idx1+1));
//					if(tmpLongitud > anchos[0]){
//						anchos[0] = tmpLongitud;
//					}
//				}else{
//					//tablaResultado.asignaValor(""+(idx1+1), idx1, 0);
//					for(int idx2 = 0; idx2 < numeroColumnas; idx2++ ){
//						String tmpPalabra = inResulSet.getString(idx2+1);
//						
//						tablaResultado.asignaValor(tmpPalabra,idx1,idx2);
//						
//						int tmpLongitud = Restriccion.getLengthString(tmpPalabra);
//						if(tmpLongitud > anchos[idx2]){
//							anchos[idx2] = tmpLongitud;
//						}
//					}
//				}
//				
//				idx1++;
//			}
//		}catch(Exception e){
//			e.printStackTrace();
//			}
//			
//		//cambia formato de tabla
//		for(int idx3 = 0; idx3 < numeroColumnas; idx3++ ){
//			tablaResultado.setAnchoCelda(idx3, anchos[idx3] + 0);
//		}
//		tablaResultado.asignarAlturaCelda(alturaCelda);
//		
//		
//		//cambio de formato de las tablas segun la opcion seleccionada
//		
//		if(!inOpcionSel){
//			JTextField miField = new JTextField();
//			miField.setEditable(false);
//			miField.setFocusable(false);
//			//miField2.setOpaque(false);
//			for(int idx4 = 0; idx4 < numeroColumnas; idx4++){
//				tablaResultado.miTabla.getColumnModel().getColumn(idx4).setCellEditor(new DefaultCellEditor(miField));
//			}
//		}
//		
//		//actualizacion de la interfaz	
//		inPanel.updateUI();
//		
//		return tablaResultado;
//
//	}													
}