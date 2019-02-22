using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net.Mail;
using System.Drawing;
using System.Net.Mime;
using System.IO;
using System.Drawing.Imaging;
using System.Data;

namespace Myp_Email.Class
{
    public class Class_email : Class_ejecutar
    {
        static string nombre = "";
        static string email = "";
        static string password = "";
        static string smtpcliente = "";
        static string puerto = "";
        static bool requiere_ssl;

        public Class_email()
        {
            //*******************************************//          
            //**** Datos del servidor de correo ********//
            //*****************************************//
            DataTable dt = new DataTable();
                DataTable dt_copy = new DataTable();
                dt = _ejecutar("select * from view_correo_envio", "2");
            //*****************************************//
            //***** var datos en las variables ********//
            //*****************************************//
            nombre = dt.Rows[0]["nombre"].ToString();
                email = dt.Rows[0]["usuario"].ToString();
                password = dt.Rows[0]["password"].ToString();
                smtpcliente = dt.Rows[0]["servidor"].ToString();
                puerto = dt.Rows[0]["puerto"].ToString();           
                if (dt.Rows[0]["requiere_ssl"].ToString() == "1")
                {
                    requiere_ssl = true;
                }
                else
                {
                    requiere_ssl = false;
                }
            //**************************************//            
        }

        private string _mes(string fecha)
        {
            DateTime date = DateTime.Parse(fecha);

            string[] meses_array = new string[] { "","ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE" };
            var mes = "";
            if (date.Month + 1 < 13) // un año nuevo           
            {
                mes = meses_array[date.Month];
            }
            else
            {
                mes = meses_array[0].ToString();
            }

            return mes;
        }

        private string _encabezado()
        {
            string contenido = "";
            contenido = "<table style=\"width:100% \"> " +
                          "<tr>" +                            
                            "<td> <img src=\"http://mypsa.com.mx/logo.png\" /> </td>" +
                            "<td  style=\"text-align: right; font-family: Times, Times New Roman, Georgia, serif; font-size: 16px; color: #333399;\">" +
                                    "RFC: MPR990906AF4" +
                                    "<br> Privada Tecnológico No. 25" +
                                    "<br> Col. Granja Nogales Sonora, C.P. 84065" +
                                    "<br> Tel: 631-314-6263, 631-314-6193" +
                           "</td>" +
                          "</tr>" +
                        "</table> <br>";

            return contenido;
        }

        private string _procesos(DataTable dtequipo, string tipo = "", string sucursal = "")
        {
            string contenido = "";

            if (tipo.ToLower() != "cliente")
            {
                contenido += " <table id=\"t02\" >" +
                   "<tr>" +
                       "<td > <strong> Recordatorio </strong> </td> " +
                       "<td > <strong> Sucursal </strong></td> " +
                   "</tr> " +
                   "<tr> " +
                       "<td>" + tipo + " </td> " +
                       "<td> " + sucursal + " </td> " +
                   "</tr> " +
                  "</table> ";
            }


            contenido += "<p style =\"text-align: left; font-family: Times, Times New Roman, Georgia, serif; font-size: 20px; color: #333634;\">Total :  " + dtequipo.Rows.Count + "</p>";


            contenido += "<table id=\"t03\"><tr>" +
                          "<th> # </th>";

            if (tipo.ToLower() == "reporte")
            {
                contenido += "<th>Cliente (Empresa / Planta)</th>" +
                "<th >Contacto (s)</th>" +
                "<th >Total de equipos</th>" +
                "<th >Correo Enviado</th>" +
                "<th >Fecha</th>";

                contenido += "</tr>";
                for (int i = 0; i < dtequipo.Rows.Count; i++)
                {
                    if ((i % 2) == 0)
                    {
                        contenido += "<tr style='background-color: #ffffff;'>";
                    }
                    else
                    {
                        contenido += "<tr style='background-color: #eeeeee ;'>";
                    }

                    var count = i + 1;
                    contenido += "" +
                         "<td>" + count + "</td>";
                    contenido += "<td >" + dtequipo.Rows[i]["cliente"] + "</td>" +
                     "<td >" + dtequipo.Rows[i]["contactos"] + "</td>" +
                     "<td >" + dtequipo.Rows[i]["total"] + "</td>" +
                     "<td >" + dtequipo.Rows[i]["enviado"] + "</td>" +
                     "<td >" + dtequipo.Rows[i]["fecha"] + "</td>";
                    contenido += "</tr>";
                }


            }
            else
            {
                if (tipo != "cliente") { contenido += "<th >Id</th>"; }

                contenido += "<th>Id equipo</th>" +
                "<th >Equipo</th>" +
                "<th >Marca</th>" +
                "<th >Modelo</th>" +
                "<th >Serie</th>";

                if (tipo == "cliente")
                {
                    contenido += "<th >Fecha de vencimiento</th>";
                }

                if (tipo != "cliente")
                {
                    contenido += "<th >Cliente</th>" +
                    "<th >Dirección</th>" +
                    "<th >Fecha de entrada</th>" +
                    "<th >Días trans.</th>";
                }

                contenido += "</tr>";
                for (int i = 0; i < dtequipo.Rows.Count; i++)
                {
                    if ((i % 2) == 0)
                    {
                        contenido += "<tr style='background-color: #ffffff;'>";
                    }
                    else
                    {
                        contenido += "<tr style='background-color: #eeeeee ;'>";
                    }
                    var count = i + 1;
                    contenido += "" +
                         "<td>" + count + "</td>";
                    if (tipo != "cliente") { contenido += "<td>" + dtequipo.Rows[i]["id"] + "</td>"; }

                    contenido += "<td >" + dtequipo.Rows[i]["id_equipo"] + "</td>" +
                     "<td >" + dtequipo.Rows[i]["descripcion"] + "</td>" +
                     "<td >" + dtequipo.Rows[i]["marca"] + "</td>" +
                     "<td >" + dtequipo.Rows[i]["modelo"] + "</td>" +
                     "<td >" + dtequipo.Rows[i]["serie"] + "</td>";

                    if (tipo == "cliente")
                    {
                        contenido += "<td >" + Convert.ToDateTime(dtequipo.Rows[i]["fecha_vencimiento"]).ToString("yyyy/MM/dd") + " </td>";
                    }

                    if (tipo != "cliente")
                    {
                        contenido += "<td >" + dtequipo.Rows[i]["cliente"] + "</td>" +
                         "<td >" + dtequipo.Rows[i]["direccion"] + "</td>" +
                         "<td >" + Convert.ToDateTime(dtequipo.Rows[i]["fecha_inicio"]).ToString("yyyy/MM/dd") + "</td>" +
                         "<td >" + dtequipo.Rows[i]["dias"] + "</td>";
                    }
                    contenido += "</tr>";
                }
            }

            contenido += "</table>";

            return contenido;
        }

        public string _infocliente(string cliente = "", string dir = "", string rfc = "", string correo = "", string nombre = "", string contacto = "",string fecharecord="")
        {


            if (String.IsNullOrEmpty(rfc) == true)
            {
                rfc = "Sin RFC";
            }

            string contenido = "";
            contenido = "<table id=\"t02\">" +
                    "<tr>" +
                        "<th> Información del cliente</ th>" +
                        "<th></th>" +
                    "</tr>" +
                    "<tr>" +
                    "<td>" +
                        "<p> <strong> Cliente:</strong> " + cliente + "</ p>" +
                        "<p> <strong> Dirección:</strong> " + dir + "</p>" +
                        "<p> <strong> RFC:</strong> " + rfc + "</p>" +
                        "<p> <strong> Contacto(s):</strong> " + correo + "</p>" +
                    "</td>" +
                        "<td><p> <strong>Contacto: " + nombre + "</strong> </p>" +
                        "<p> De acuerdo a nuestros registros, la calibración del equipo listado a continuación vencerá "+
                        "el próximo mes de <strong> "+ _mes(fecharecord) + "</strong>  del   <strong>" + DateTime.Now.Year.ToString() + "</strong>. </p>" +
                        "<p> <strong>Nota:</strong> Este correo se envía de manera automática, favor de responder a la dirección de correo siguiente: <br> <a>" + contacto + " </a> </p>" +
                    "</td>" +
                   "</tr>" +
                   "</table>";
            return contenido;
        }

        public string _enviar(DataTable dtequipo, string tipo = "", string correo = "", string info_cliente = "", string sucursal = "")
        {
            string toadd = "";
            string bccadd = "";            
            bool html = false;            
            string bodyhtml = "";
            string asunto = "";
            toadd = correo;            
            //oMsg.Bcc.Add(new MailAddress("test@mypsa.com.mx"));
            if (tipo=="cliente" && sucursal=="2")
            {
                bccadd="lromero@mypsa.mx";
            }
            asunto = "Recordatorio de calibración";
            // Add HTML and text body
            html = true;            
            //Nuevo stilo
            bodyhtml += "<!DOCTYPE html> <html>" +
                    "<head>" +
                    "<title>Recordatorio</title>" +
                    "<style> " +
                        "table {width:100%;}" +
                        "table, th, td {border: 1px; border-collapse: collapse; }" +
                        "table#t02 td {width: 50%; text-align: center; background-color: #eeeeee; padding: 5px; font-family:Times, Times New Roman, Georgia, serif; font-size:14px;}" +
                        "table#t03 td {text-align: left; font-family:Times, Times New Roman, Georgia, serif; font-size:12px; padding: 5px; border-bottom: 1px solid #D6D6D6; }" +
                        "table#t03 th {background-color: #009AD6; color: white; text-align: left; font-size:14px; padding: 5px; border: 0px}" +
                    "</style>" +
                    "</head>";
            bodyhtml += "<body>";
            //Fin           
            bodyhtml += _encabezado();
            if (tipo != "cliente")
            {
                bodyhtml += _procesos(dtequipo, tipo, sucursal);
            }
            else
            {
                bodyhtml += info_cliente;
                bodyhtml += _procesos(dtequipo, tipo);
            }
            bodyhtml += "</body></html>";          

            return _enviodecorreo(asunto,toadd,bccadd, html, bodyhtml);
        }

        public string _enviartemp(string asunto = "", string correo = "", string body = "", bool htmlbody = false) //Template
        {
            string toadd = "";
            string bccadd = "";            
            bool html = false;
            string bodyhtml = "";            
            bccadd = correo;            
            // Add HTML and text body
            html = htmlbody;
            bodyhtml = body;

            return _enviodecorreo(asunto, toadd, bccadd, html, bodyhtml);
        }

        public string _enviodecorreo(string asunto,string toadd, string bccadd, bool html, string body)
        {
            //Creamos un nuevo Objeto de Mensaje
            MailMessage oMsg = new MailMessage();            
            //Desde (correo electronico del que enviamos)
            oMsg.From = new MailAddress(email, nombre);
            //Asunto del correo
            oMsg.Subject = asunto;
            //correo de copia
            if (!String.IsNullOrEmpty(toadd))
            {
                oMsg.To.Add(toadd);
            }
            //Correo de copias en oculto
            if (!String.IsNullOrEmpty(bccadd))
            {
                oMsg.Bcc.Add(bccadd);
            }
            oMsg.Bcc.Add(new MailAddress("it@mypsa.mx", "copia de envio"));
            // Activar/Desactivar el correo en formato HTML
            oMsg.IsBodyHtml = html;
            //Cuerpo del correo
            oMsg.Body = body;
            try
            {
                using (SmtpClient smtp = new SmtpClient(smtpcliente))
                {
                    smtp.Port = Int32.Parse(puerto);
                    smtp.EnableSsl = requiere_ssl;                    
                    smtp.Credentials = new System.Net.NetworkCredential(email, password);
                    smtp.Send(oMsg);
                }
                return "Exitoso";
            }
            catch (Exception ex)
            {
                ex.ToString();
                return "Error";
            }           
        }
    }
}
