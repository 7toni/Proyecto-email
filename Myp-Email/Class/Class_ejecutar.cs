using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MySql.Data.MySqlClient;


namespace Myp_Email.Class
{
    public class Class_ejecutar
    {
        DataTable dtequipos = new DataTable();
        public Class_ejecutar()
        {
            //
        }

        public DataTable _ejecutar(string query = "", string tabla = "")
        {

            dtequipos.Clear();
            Class_conexion class_cnx = new Class_conexion();

            MySqlConnection cnx = new MySqlConnection(class_cnx._conexion(tabla));
            cnx.Open();
            MySqlCommand cmd = new MySqlCommand(query, cnx);
            MySqlDataAdapter da = new MySqlDataAdapter(cmd);
            DataSet ds = new DataSet();
            da.Fill(ds);
            dtequipos = ds.Tables[0];
            cnx.Close();
            return dtequipos;
        }

        public DataTable _select(string proc = "", string suc = "", string fecha = "")
        {
            DataTable dt = new DataTable();
            try
            {
                string strQuery = _querys(proc, suc, fecha);
                dt = _ejecutar(strQuery);
                return dt;
            }
            catch (Exception)
            {
                return dt;
            }
        }

        public string _querys(string opcion = "", string suc = "", string fecha = "")
        {
            string consulta = "";
            switch (opcion)
            {
                case "calibracion": //opción para saber los equipos en proceso de calibración
                    consulta = "SELECT id,alias as id_equipo,descripcion,marca,modelo,serie, concat( empresa ,' ',if((planta= 'Planta1' or planta= 'Planta 1'),'',concat(' (',planta,')'))) as Cliente, direccion,usuarios_calibracion_id as id_tecnico,calibrado_por as tecnico,fecha_hoja_entrada as fecha_inicio  FROM view_informes_" + suc + " where proceso= 1 and fecha_hoja_entrada is not null order by usuarios_calibracion_id desc";
                    break;
                case "salida"://opción para saber los equipos en proceso de salida
                    consulta = "SELECT id,alias as id_equipo,descripcion,marca,modelo,serie, concat( empresa ,' ',if((planta= 'Planta1' or planta= 'Planta 1'),'',concat(' (',planta,')'))) as Cliente, direccion,fecha_hoja_entrada as fecha_inicio  FROM view_informes_" + suc + " where proceso= 2 and fecha_hoja_entrada is not null order by id desc";
                    break;
                case "facturacion"://opción para saber los equipos en proceso de facturación
                    consulta = "SELECT id,alias as id_equipo,descripcion,marca,modelo,serie, concat( empresa ,' ',if((planta= 'Planta1' or planta= 'Planta 1'),'',concat(' (',planta,')'))) as Cliente, direccion,fecha_hoja_entrada as fecha_inicio  FROM view_informes_" + suc + " where proceso= 3 and fecha_hoja_entrada is not null order by id desc";
                    break;
                case "correo_tec":
                    consulta = "SELECT email FROM view_usuarios where id=" + int.Parse(suc) + " and activo='si'"; // Id del técnico
                    break;
                case "correo_cliente":
                    consulta = "SELECT email FROM view_usuarios where plantas_id=" + int.Parse(suc) + " and (roles_id=10002 || roles_id=10004  || roles_id=10005) and activo='si' and email not like '%cliente%' "; // Id del cliente
                    break;
                case "clientes":                   
                    consulta = "SELECT id,alias as id_equipo,descripcion,marca,modelo,serie, plantas_id as id_cliente, concat( empresa ,' ',if((planta= 'Planta1' or planta= 'Planta 1'),'',concat(' (',planta,')'))) as cliente, direccion,rfc,fecha_vencimiento as fecha_vencimiento FROM view_informes_" + suc + " where  periodo_calibracion> 0  and  fecha_vencimiento between ('" + fecha + "') and (date_add('" + fecha + "', interval 1 month)) and month(fecha_vencimiento)= month(date_add('" + fecha + "', interval 1 month)) and calibraciones_id != 3 and plantas_id is not null order by id_cliente, fecha_vencimiento asc"; // query para calcular todos los equipos vencidos del siguiente mes                                    
                    break;
                case "correo_calibracion":
                    consulta = "SELECT * FROM view_" + opcion + " where id_sucursal=" + int.Parse(suc);
                    break;
                case "correo_salida":
                    consulta = "SELECT * FROM view_" + opcion + " where id_sucursal=" + int.Parse(suc);
                    break;
                case "correo_facturacion":
                    consulta = "SELECT * FROM view_" + opcion + " where id_sucursal=" + int.Parse(suc);
                    break;
                case "correo_cotizacion":
                    consulta = "SELECT * FROM view_" + opcion + " where id_sucursal=" + int.Parse(suc);
                    break;
                case "correo_reporte":
                    consulta = "SELECT * FROM view_" + opcion + " where id_sucursal=" + int.Parse(suc);
                    break;
                default:
                    break;

            }
            return consulta;
        }

        public string _add(string tabla = "", string values = "")
        {
            try
            {
                string celdas = "";
                if (tabla == "registro") { celdas = "idoperacion,idsucursal,estado,detalle"; }
                else { celdas = "nombre,correo,id_sucursal"; }

                string consulta = "insert into " + tabla + "(" + celdas + ") VALUES (" + values + ");";
                _insert(consulta);
                return "Exitoso";
            }
            catch (Exception ex)
            {
                return ex.ToString();
            }
        }

        public string _update(string tabla = "", string values = "")
        {
            try
            {
                string consulta = " UPDATE " + tabla + " SET " + values;
                _insert(consulta);
                return "Exitoso";
            }
            catch (Exception ex)
            {
                return ex.ToString();
            }
        }

        public string _delete(string tabla = "", string values = "")
        {
            try
            {
                string consulta = " DELETE FROM " + tabla + " WHERE " + values;
                _insert(consulta);
                return "Exitoso";
            }
            catch (Exception ex)
            {
                return ex.ToString();
            }
        }

        public void _insert(string query = "")
        {
            Class_conexion class_cnx = new Class_conexion();
            MySqlConnection cnx = new MySqlConnection(class_cnx._conexion("2"));
            cnx.Open();
            MySqlCommand cmd = new MySqlCommand(query, cnx);
            cmd.ExecuteNonQuery();
            cnx.Close();
        }
    }
}
