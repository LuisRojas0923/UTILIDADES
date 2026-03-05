/*
 Proyecto Sistema Integrado de Produccion (SOLID) THC
 
 Clase 				ParametroInforme
 Archivo			ParametroInforme.java
  
 Fecha Creación 	6 de Junio de 2008
 Autor  			Marlon Ramirez Jaime 1.130.637.509
 email 				marlonrj@gmail.com
*/

import javax.swing.*;
import java.awt.*;
import java.awt.event.*;
import java.util.*;
import javax.swing.*;
import java.sql.*;

/**
 * Esta clase guarda la informacion de un parametro de informe
 *
 */
public class ParametroInforme {
	//atributos
	/** codigo de ParametroInforme */
	public String codigo;
	
	/** parametro de ParametroInforme */
	public String parametro;

	/** interface con el manejador de base de datos */
	public ConexionDatos miConexionDatos = null;
		
	
	//termina sql
	
	//metodos
	/**
	 * constructor por defecto
	 *
	 * @param in_codigo -codigo de ParametroInforme
	 * @param inRepresentacion -parametro de ParametroInforme
	 * @param in_conexionDatos -interfaz con base de datos que se utilizara
	 *
	 */
	ParametroInforme(	String inCodigo, 
						String inParametro,
						ConexionDatos inConexionDatos){
		
		codigo = inCodigo;		
		parametro = inParametro;
		
		//para mejorar el rendimiento se referencia la conexion
		//conexion base de datos
		miConexionDatos = inConexionDatos;
		//
	}
	/**
	 * constructor para consultar
	 *
	 * @param inCodigo -codigo de la ParametroInforme
	 */
	ParametroInforme(		String inCodigo,
							ConexionDatos inConexionDatos ){
		codigo = inCodigo;
		
		//para mejorar el rendimiento se referencia la conexion
		//conexion base de datos
		miConexionDatos = inConexionDatos;
		//
	}
	/**
	 * guarda los registros en la base de datos
	 * 
	 * @return -devuelve 1 si la operacion fue exitosa, 0 si no lo fue
	 *
	 */
	public int guardar(){
		return miConexionDatos.funcionActualizar2("insert into parametroInforme values("
													+codigo+",'"	
													+parametro+"')");
	}
	/**
	 * carga los registros de ParametroInforme de la base de datos
	 * 
	 * @param atributo -parametro del campo comparado
	 * @param referencia -palabra con la que se compara
	 *
	 */
	public void cargar(String inAtributo, String inReferencia){
		ResultSet resultSet;
		int indice = 0;
		resultSet = miConexionDatos.funcionConsultar("select * from parametroInforme where "+inAtributo+" = '"+inReferencia+"'");
		try{
			while(resultSet.next()){
				indice++;
				codigo = (String)resultSet.getString("codigo");
				parametro = (String)resultSet.getString("parametro");
			}
			if (indice == 0){
				JOptionPane.showMessageDialog(null,"El parametro de informe no existe");
			}
		}catch(Exception e){
			e.printStackTrace();
		}
	}
	/**
	 * determina la cantidad de tuplas existentes con la caracteristica especifica
	 * 
	 * @param atributo -parametro del campo comparado
	 * @param referencia -palabra con la que se compara
	 *
	 * @return -retorna el numero de tuplas que coinciden con la consulta
	 *
	 */
	public static int contar(String inAtributo, String inReferencia, ConexionDatos inConexionDatos){
		ResultSet resultSet;
		int indice = 0;
		resultSet = inConexionDatos.funcionConsultar("select * from parametroInforme where "+inAtributo+" = '"+inReferencia+"'");
		try{
			while(resultSet.next()){
				indice++;
			}
		}catch(Exception e){
			e.printStackTrace();
		}
		return indice;
	}
	/**
	 * actualiza los registros en la base de datos de la ParametroInforme que tenga el codigo dado
	 * 
	 * @return -devuelve 1 si la operacion fue exitosa, 0 si no lo fue
	 *
	 */
	public int actualizar(){
			return miConexionDatos.funcionActualizar2("update parametroInforme set "
															+" parametro = '"+parametro+"' "
															
															+" where "
															+" codigo = '"+codigo+"'");
	}
	/**
	 * elimina los registros del ParametroInforme de la base de datos 
	 * 
	 * @return -devuelve 1 si la operacion fue exitosa, 0 si no lo fue
	 *
	 */
	public int eliminar(){
		return miConexionDatos.funcionActualizar("delete from parametroInforme "
						+" where "
						+" codigo = '"+codigo+"'");
	}
}