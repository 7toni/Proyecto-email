using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Myp_Email.Class
{

    public class Class_debug : Class_ejecutar
    {
        Class.Class_email correo = new Class.Class_email();

        string nom_cliente = "", dir_cliente = "", rfc_cliente = "", nombre = "", contacto = "";

        public Class_debug()
        {
            //
        }

        public DataTable _editar(DataTable dtequipo)
        {
            dtequipo.Columns.Add("dias", typeof(int));
            DataRow renglon;
            renglon = dtequipo.NewRow();
            int filas = dtequipo.Rows.Count;
            for (int i = 0; i < filas; i++)
            {
                DateTime fecha = Convert.ToDateTime(dtequipo.Rows[i]["fecha_inicio"]);
                var dias = _diastrans(fecha);
                dtequipo.Rows[i]["dias"] = dias;
            }
            return dtequipo;
        }

        private int _diastrans(DateTime fecha)
        {
            var dias = 0;
            var diastrans = 0;
            string fecha2 = fecha.ToString("dd/MM/yyyy");
            //DateTime olddate = Convert.ToDateTime(fecha2);
            DateTime olddate = DateTime.ParseExact(fecha2, "dd/MM/yyyy", null);
            DateTime newdate = DateTime.Now;
            var diashabiles = (newdate - olddate).TotalDays;
            var finsemana = olddate.DayOfWeek - newdate.DayOfWeek;
            double calcBusinessDays = ((diashabiles) * 5 - (finsemana) * 2) / 7;

            if ((int)newdate.DayOfWeek == 6) calcBusinessDays--;
            if ((int)olddate.DayOfWeek == 0) calcBusinessDays--;

            diastrans = (int)calcBusinessDays;
            if (diastrans < 1) // dias trancsurridos  menores que 3 se asignara 0 días
            {
                dias = 0;
            }
            else
            {
                dias = diastrans;
            }
            return dias;
        }

        public string _tecnico(DataTable dtequipo,string sucursal="")
        {
            string result = "";

            DataTable dt_tecnico = new DataTable();
            dt_tecnico.Columns.Add("id", typeof(string));
            dt_tecnico.Columns.Add("id_equipo", typeof(string));
            dt_tecnico.Columns.Add("descripcion", typeof(string));
            dt_tecnico.Columns.Add("marca", typeof(string));
            dt_tecnico.Columns.Add("modelo", typeof(string));
            dt_tecnico.Columns.Add("serie", typeof(string));
            dt_tecnico.Columns.Add("cliente", typeof(string));
            dt_tecnico.Columns.Add("direccion", typeof(string));
            dt_tecnico.Columns.Add("id_tecnico", typeof(string));
            dt_tecnico.Columns.Add("fecha_inicio", typeof(string));
            dt_tecnico.Columns.Add("dias", typeof(string));
            DataRow row;

            //int filas = dtequipo.Rows.Count;
            int i;
            string nom_tecnico = "";
            try
            {
                for (i = 0; i < dtequipo.Rows.Count; i++)
                {
                    var igual = 0;
                    var id_tecnico = int.Parse(dtequipo.Rows[i]["id_tecnico"].ToString());
                    if (i == 0) // 
                    {
                        nom_tecnico = dtequipo.Rows[i]["tecnico"].ToString();
                        row = dt_tecnico.NewRow();
                        row["id"] = dtequipo.Rows[i]["id"];
                        row["id_equipo"] = dtequipo.Rows[i]["id_equipo"];
                        row["descripcion"] = dtequipo.Rows[i]["descripcion"];
                        row["marca"] = dtequipo.Rows[i]["marca"];
                        row["modelo"] = dtequipo.Rows[i]["modelo"];
                        row["serie"] = dtequipo.Rows[i]["serie"];
                        row["cliente"] = dtequipo.Rows[i]["cliente"];
                        row["direccion"] = dtequipo.Rows[i]["direccion"];
                        row["id_tecnico"] = dtequipo.Rows[i]["id_tecnico"];
                        row["fecha_inicio"] = dtequipo.Rows[i]["fecha_inicio"];
                        row["dias"] = dtequipo.Rows[i]["dias"];
                        dt_tecnico.Rows.Add(row);

                    }
                    else
                    {
                        var id_tecnicotemp = int.Parse(dt_tecnico.Rows[0]["id_tecnico"].ToString());
                        if (id_tecnico == id_tecnicotemp) //i > 0
                        {
                            row = dt_tecnico.NewRow();
                            row["id"] = dtequipo.Rows[i]["id"];
                            row["id_equipo"] = dtequipo.Rows[i]["id_equipo"];
                            row["descripcion"] = dtequipo.Rows[i]["descripcion"];
                            row["marca"] = dtequipo.Rows[i]["marca"];
                            row["modelo"] = dtequipo.Rows[i]["modelo"];
                            row["serie"] = dtequipo.Rows[i]["serie"];
                            row["cliente"] = dtequipo.Rows[i]["cliente"];
                            row["direccion"] = dtequipo.Rows[i]["direccion"];
                            row["id_tecnico"] = dtequipo.Rows[i]["id_tecnico"];
                            row["fecha_inicio"] = dtequipo.Rows[i]["fecha_inicio"];
                            row["dias"] = dtequipo.Rows[i]["dias"];
                            dt_tecnico.Rows.Add(row);
                        }
                        else
                        {
                            for (int j = (i - 1); j >= 0; j--)
                            {
                                DataRow rowd = dtequipo.Rows[j];
                                dtequipo.Rows.Remove(rowd);
                            }
                            var id_usuario = int.Parse(dt_tecnico.Rows[0]["id_tecnico"].ToString()); //Obtenemos el id del tecnico
                            DataTable temp = new DataTable();
                            temp = _select("correo_tec", id_usuario.ToString()); // Select el correo del tecnico
                            if (temp.Rows.Count>0)
                            {
                                var correo_tec = temp.Rows[0]["email"].ToString();
                                correo._enviar(dt_tecnico, nom_tecnico, correo_tec, "", sucursal);                                
                            }
                            dt_tecnico.Clear();
                            i = -1;
                        }

                    }
                    igual = i + 1;
                    if (dtequipo.Rows.Count == igual)
                    {
                        for (int j = i; j >= 0; j--)
                        {
                            DataRow rowd = dtequipo.Rows[j];
                            dtequipo.Rows.Remove(rowd);
                        }
                        var id_usuario = int.Parse(dt_tecnico.Rows[0]["id_tecnico"].ToString()); //Obtenemos el id del tecnico
                        DataTable temp = new DataTable();
                        temp = _select("correo_tec", id_usuario.ToString()); // Select el correo del tecnico                       
                        if (temp.Rows.Count > 0)
                        {
                            var correo_tec = temp.Rows[0]["email"].ToString();
                            correo._enviar(dt_tecnico, nom_tecnico, correo_tec, "", sucursal);                            
                        }
                        dt_tecnico.Clear();
                        i = -1;
                    }
                }
                result = "exitoso";
            }
            catch (Exception ex)
            {
                result= ex.ToString();
            }

            return result;
        }

        public DataTable _clientes(DataTable dtequipo, int suc = 0)
        {
            DataTable dt_cliente = new DataTable();
            dt_cliente.Columns.Add("id_equipo", typeof(string));
            dt_cliente.Columns.Add("descripcion", typeof(string));
            dt_cliente.Columns.Add("marca", typeof(string));
            dt_cliente.Columns.Add("modelo", typeof(string));
            dt_cliente.Columns.Add("serie", typeof(string));
            dt_cliente.Columns.Add("fecha_vencimiento", typeof(string));
            dt_cliente.Columns.Add("id_cliente", typeof(string));          

            DataTable dt_repor_cliente = new DataTable();           
            dt_repor_cliente.Columns.Add("cliente", typeof(string));
            dt_repor_cliente.Columns.Add("contactos", typeof(string));
            dt_repor_cliente.Columns.Add("total", typeof(string));
            dt_repor_cliente.Columns.Add("enviado", typeof(string));
            dt_repor_cliente.Columns.Add("fecha", typeof(string)); 
            
            DataRow row;

            dt_cliente.Clear();
            dt_repor_cliente.Clear();
            string fecha="";

            //int filas = dtequipo.Rows.Count;
            int i;            
            DataTable dt_email = new DataTable();
            dt_email = _ejecutar(_querys("correo_cotizacion", suc.ToString()), "2"); //obtengo el query, deespues el dt
            if (dt_email.Rows.Count > 0)
            {
                nombre = dt_email.Rows[0]["nombre"].ToString();
                contacto = dt_email.Rows[0]["email"].ToString();
            }
            else {
                nombre = "";
                contacto = "";
            }
           

            for (i = 0; i < dtequipo.Rows.Count; i++)
            {
                var igual = 0;
              var id_cliente = int.Parse(dtequipo.Rows[i]["id_cliente"].ToString());
                if (i == 0) // 
                {
                    nom_cliente = dtequipo.Rows[i]["cliente"].ToString();
                    dir_cliente = dtequipo.Rows[i]["direccion"].ToString();
                    rfc_cliente = dtequipo.Rows[i]["rfc"].ToString();
                    fecha = dtequipo.Rows[i]["fecha_vencimiento"].ToString();

                    row = dt_cliente.NewRow();
                    row["id_equipo"] = dtequipo.Rows[i]["id_equipo"];
                    row["descripcion"] = dtequipo.Rows[i]["descripcion"];
                    row["marca"] = dtequipo.Rows[i]["marca"];
                    row["modelo"] = dtequipo.Rows[i]["modelo"];
                    row["serie"] = dtequipo.Rows[i]["serie"];
                    row["fecha_vencimiento"] = dtequipo.Rows[i]["fecha_vencimiento"];
                    row["id_cliente"] = dtequipo.Rows[i]["id_cliente"];
                    dt_cliente.Rows.Add(row);
                }
                else
                {
                    var id_clientetemp = int.Parse(dt_cliente.Rows[0]["id_cliente"].ToString());
                    if (id_cliente == id_clientetemp)
                    {
                        row = dt_cliente.NewRow();
                        row["id_equipo"] = dtequipo.Rows[i]["id_equipo"];
                        row["descripcion"] = dtequipo.Rows[i]["descripcion"];
                        row["marca"] = dtequipo.Rows[i]["marca"];
                        row["modelo"] = dtequipo.Rows[i]["modelo"];
                        row["serie"] = dtequipo.Rows[i]["serie"];
                        row["fecha_vencimiento"] = dtequipo.Rows[i]["fecha_vencimiento"];
                        row["id_cliente"] = dtequipo.Rows[i]["id_cliente"];
                        dt_cliente.Rows.Add(row);
                    }
                    else //Cuando se detecta un cambio de cliente se procesa
                    {
                        for (int j = (i - 1); j >= 0; j--)
                        {
                            DataRow rowd = dtequipo.Rows[j];
                            dtequipo.Rows.Remove(rowd);
                        }
                        var cliente = int.Parse(dt_cliente.Rows[0]["id_cliente"].ToString()); //Obtenemos el id del cliente
                        DataTable temp = new DataTable();
                        temp = _select("correo_cliente", cliente.ToString()); // Select el correo del cliente
                        string lista_email = _listaemails(temp);
                        var email_cliente="";
                        if (String.IsNullOrEmpty(lista_email))
                        {
                            email_cliente = "Sin identificación";
                            lista_email = contacto;
                        }
                        else {
                            email_cliente = lista_email;                           
                        }
                        string html = correo._infocliente(nom_cliente, dir_cliente, rfc_cliente, email_cliente, nombre, contacto, fecha);
                        var confirmacion = correo._enviar(dt_cliente, "cliente", lista_email, html,suc.ToString());
                        //Inicio de
                        row = dt_repor_cliente.NewRow();
                        row["cliente"] = nom_cliente;
                        row["contactos"] = email_cliente;
                        row["total"] = dt_cliente.Rows.Count;
                        row["enviado"] = confirmacion;
                        row["fecha"] = DateTime.Now.ToString();
                        dt_repor_cliente.Rows.Add(row);
                        // Fin de 
                        dt_cliente.Clear();
                        i = -1;
                    }

                }
                igual = i +1;
                if (dtequipo.Rows.Count == igual) //38 / i=37
                {                   
                    for (int j = i; j >= 0; j--)
                    {
                        DataRow rowd = dtequipo.Rows[j];
                        dtequipo.Rows.Remove(rowd);                    
                    }
                    var cliente = int.Parse(dt_cliente.Rows[0]["id_cliente"].ToString()); //Obtenemos el id del tecnico
                    DataTable temp = new DataTable();
                    temp = _select("correo_cliente", cliente.ToString()); // Select el correo del tecnico
                    string lista_email = _listaemails(temp);
                    var email_cliente = "";
                    if (String.IsNullOrEmpty(lista_email))
                    {
                        email_cliente = "Sin identificación";
                        lista_email = contacto;
                    }
                    else
                    {
                        email_cliente = lista_email;
                    }
                    string html = correo._infocliente(nom_cliente, dir_cliente, rfc_cliente, email_cliente, nombre, contacto,fecha);
                    var confirmacion =correo._enviar(dt_cliente, "cliente", lista_email, html, suc.ToString());
                    //Inicio de
                    row = dt_repor_cliente.NewRow();
                    row["cliente"] = nom_cliente;
                    row["contactos"] = email_cliente;
                    row["total"] = dt_cliente.Rows.Count;
                    row["enviado"] = confirmacion;
                    row["fecha"] = DateTime.Now.ToString();
                    dt_repor_cliente.Rows.Add(row);
                    // Fin de 
                    dt_cliente.Clear();
                    i = -1;
                }
            }
            return dt_repor_cliente;
        }

        public string _listaemails(DataTable dt)
        {
            string contactos = "";
            foreach (DataRow fila in dt.Rows)
            {
                contactos += fila["email"].ToString() + ", ";
            }
            if (contactos.Length > 0) { contactos = contactos.Substring(0, (contactos.Length) - 2); }            
            return contactos;
        }


    }



}
