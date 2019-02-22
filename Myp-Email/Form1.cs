using MetroFramework;
using MetroFramework.Forms;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Myp_Email
{
    public partial class Form1 : MetroFramework.Forms.MetroForm
    {
        DataTable dt = new DataTable();
        DataTable dt2 = new DataTable();
        Class.Class_debug her_depurar_ejecutar = new Class.Class_debug();
        Class.Class_email correo = new Class.Class_email();
        bool filecsv = false;


        string[] ArrSucur = { "n", "h", "g" };
        string[] ArrNomSucur = { "Nogales", "Hermosillo", "Guaymas" };
        DateTime fechainicial, diaultimomes;
        string horainicial2 = "07:35";
        int hora = 8, minuto = 15, segundo = 0;
        bool posponer = false;
        public Form1()
        {

            Thread t = new Thread(new ThreadStart(Loading));
            t.Start();
            InitializeComponent();
            for (int i = 0; i < 1000; i++)
            {
                Thread.Sleep(10);
            }
            t.Abort();
            this.StyleManager = this.metroStyleManager1;
            this.BorderStyle = MetroFormBorderStyle.FixedSingle;
            this.ShadowType = MetroFormShadowType.AeroShadow;
            timer_reloj.Start();
            timer_checador.Interval = (1000 * 60); // cada  minuto
            timer_checador.Start();
            check_n1.Checked = true;
            check_h1.Checked = true;
            check_g1.Checked = false;
            metroCheckBox5.Checked = true;
            MaximizeBox = false;
            metroTextBox_correos.Enabled = false;
            metroButtonFileCsv.Enabled = false;
            metroComboproceso.SelectedIndex = 0;
            DateTime hoy = DateTime.Now;
            fechainicial = new DateTime(hoy.Year, hoy.Month, hoy.Day, 7, 55, 0); // Parametro de 24 horas 
        }
        void Loading()
        {
            frmMain frm = new frmMain();
            Application.Run(frm);
        }

        private void _procesos(string proceso = "", string desc = "", string correotep = "", string fecha="")
        {
            string detalle = "";
            bool[] Arrcheck = { check_n1.Checked, check_h1.Checked, check_g1.Checked };
            for (int i = 0; i < 3; i++)
            {
                int y = i + 1;
                dt.Clear();
                if (Arrcheck[i] == true)
                {
                    if (proceso == "calib_tec")
                    {
                        detalle = "Inicio: " + DateTime.Now.ToString() + " - Recordatorio de técnicos - Sucursal: " + ArrNomSucur[i];
                        DataTable dt_copy = new DataTable();
                        dt = her_depurar_ejecutar._editar(her_depurar_ejecutar._select("calibracion", ArrSucur[i]));
                        dt_copy = dt.Copy();
                        string result = her_depurar_ejecutar._tecnico(dt_copy, ArrNomSucur[i]);
                        if (result == "exitoso")
                        {
                            detalle += " - Fin: " + DateTime.Now.ToString() + " - Operación terminada exitosamente";
                            her_depurar_ejecutar._add("registro", "1 ," + y + ", 'exitoso','" + detalle + "'");
                        }
                        else
                        {
                            detalle += " - Fin: " + DateTime.Now.ToString() + " - Operación fallida";
                            her_depurar_ejecutar._add("registro", "1 ," + y + ", 'fallido','" + detalle + "'");
                        }

                    }
                    else if (proceso == "clientes")
                    {
                        detalle = "Inicio: " + DateTime.Now.ToString() + " - Recordatorio de clientes - Sucursal: " + ArrNomSucur[i];
                        try
                        {
                            DataTable dt_tempreport = new DataTable();
                            DataTable dt_copy = new DataTable();
                            dt = her_depurar_ejecutar._select(proceso, ArrSucur[i], fecha);
                            dt_copy = dt.Copy();
                            dt_tempreport = her_depurar_ejecutar._clientes(dt_copy, y);
                            // Enviar reporte, REvisar                            
                            DataTable dt_email = new DataTable();
                            dt_email = her_depurar_ejecutar._ejecutar(her_depurar_ejecutar._querys("correo_reporte", y.ToString()), "2"); //obtengo el query, deespues el dt
                            string listaemail = her_depurar_ejecutar._listaemails(dt_email); // Obtengo la lista de correos
                            correo._enviar(dt_tempreport, "Reporte", listaemail, "", ArrNomSucur[i]);
                            // Fin
                            detalle += " - Fin: " + DateTime.Now.ToString() + " - Operación terminada exitosamente";
                            her_depurar_ejecutar._add("registro", "5," + y + ", 'exitoso','" + detalle + "'");

                        }
                        catch (Exception)
                        {
                            detalle += " Fin: " + DateTime.Now.ToString() + " - Operación fallida";
                            her_depurar_ejecutar._add("registro", "5," + y + ", 'fallido','" + detalle + "'");
                        }
                    }
                    else
                    {
                        detalle = "Inicio: " + DateTime.Now.ToString() + " - Recordatorio de " + proceso + " - Sucursal: " + ArrNomSucur[i];
                        var operacion = 0;
                        if (proceso == "calibracion")
                            operacion = 2;
                        if (proceso == "salida")
                            operacion = 3;
                        if (proceso == "facturacion")
                            operacion = 4;

                        try
                        {
                            DataTable dt_copy = new DataTable();
                            dt = her_depurar_ejecutar._editar(her_depurar_ejecutar._select(proceso, ArrSucur[i]));
                            dt_copy = dt.Copy();
                            //Opcion para enviar correo de manera particular
                            string listaemail = "";
                            if (String.IsNullOrEmpty(correotep))
                            {
                                DataTable dt_email = new DataTable();
                                dt_email = her_depurar_ejecutar._ejecutar(her_depurar_ejecutar._querys("correo_" + proceso, y.ToString()), "2"); //obtengo el query, deespues el dt
                                listaemail = her_depurar_ejecutar._listaemails(dt_email); // Obtengo la lista de correos
                            }
                            else
                            {
                                listaemail = correotep;
                            }
                            correo._enviar(dt_copy, desc, listaemail, "", ArrNomSucur[i]);
                            detalle += " - Fin: " + DateTime.Now.ToString() + " - Operación terminada exitosamente";

                            her_depurar_ejecutar._add("registro", "" + operacion + "," + y + ", 'exitoso','" + detalle + "'");
                        }
                        catch (Exception)
                        {
                            detalle += " - Fin: " + DateTime.Now.ToString() + " - Operación fallida";
                            her_depurar_ejecutar._add("registro", "" + operacion + "," + y + ", 'fallido','" + detalle + "'");
                        }
                    }
                }
            }
        }

        private void calibraciónToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                _procesos("calibracion", "Calibración");
                MetroMessageBox.Show(this, "Se envio Exitosamente!", "Mensaje de notificación", MessageBoxButtons.OK, MessageBoxIcon.Question);
            }
            catch (Exception)
            {
                MetroMessageBox.Show(this, "Error al enviar el correo, verificar datos. Si persite el error reportarlo al administrador!", "Mensaje de notificación", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }
        }

        private void salidaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                _procesos("salida", "Salida");
                MetroMessageBox.Show(this, "Se envio Exitosamente!", "Mensaje de notificación", MessageBoxButtons.OK, MessageBoxIcon.Question);
            }
            catch (Exception)
            {
                MetroMessageBox.Show(this, "Error al enviar el correo, verificar datos. Si persite el error reportarlo al administrador!", "Mensaje de notificación", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }

        }

        private void facturaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                _procesos("facturacion", "Facturación");
                MetroMessageBox.Show(this, "Se envio Exitosamente!", "Mensaje de notificación", MessageBoxButtons.OK, MessageBoxIcon.Question);
            }
            catch (Exception)
            {
                MetroMessageBox.Show(this, "Error al enviar el correo, verificar datos. Si persite el error reportarlo al administrador!", "Mensaje de notificación", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }
        }

        private void calibracióntécnicoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                _procesos("calib_tec");
                MetroMessageBox.Show(this, "Se envio Exitosamente!", "Mensaje de notificación", MessageBoxButtons.OK, MessageBoxIcon.Question);
            }
            catch (Exception)
            {
                MetroMessageBox.Show(this, "Error al enviar el correo, verificar datos. Si persite el error reportarlo al administrador!", "Mensaje de notificación", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }
        }

        private void clientesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                DateTime hoy = DateTime.Today;
                _procesos("clientes", "", "", hoy.ToString("yyyy-MM-dd"));
                MetroMessageBox.Show(this, "Se envio Exitosamente!", "Mensaje de notificación", MessageBoxButtons.OK, MessageBoxIcon.Question);
            }
            catch (Exception)
            {
                MetroMessageBox.Show(this, "Error al enviar el correo, verificar datos. Si persite el error reportarlo al administrador!", "Mensaje de notificación", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }
        }

        private void timer_reloj_Tick(object sender, EventArgs e)
        {
            DateTime hoy = DateTime.Now;

            metrotimer.Text = DateTime.Now.ToLongTimeString();
            if (AccesoInternet())
                radio_online.Checked = true;
            else
                radio_online.Checked = false;

            Fecha.Text = hoy.ToString("dd-MMM-yyyy");
            TimeSpan ts1 = new TimeSpan(fechainicial.Hour, fechainicial.Minute, fechainicial.Second);
            TimeSpan ts2 = new TimeSpan(hoy.Hour, hoy.Minute, hoy.Second);

            int result = TimeSpan.Compare(ts2, ts1); // Compara si es antes, igual o despues de la hora (-1,0,1)
            if (result == 0) //Envia recordatorios programado
            {
                /*Verificar si hay internet, si no posponerlo */
                if (AccesoInternet())
                {
                    if (hoy.DayOfWeek != DayOfWeek.Saturday && hoy.DayOfWeek != DayOfWeek.Sunday)
                    {
                        _procesos("calib_tec");
                        Thread.Sleep(10000);
                        _procesos("calibracion", "Calibración");
                        Thread.Sleep(10000);
                        _procesos("salida", "Salida");
                        Thread.Sleep(10000);
                        _procesos("facturacion", "Facturación");
                    }
                }
                else
                {
                    fechainicial = new DateTime(hoy.Year, hoy.Month, hoy.Day, hoy.Hour, hoy.Minute, 0).AddMinutes(5); // Parametro de 24 horas  
                }

            }

        }

        private void timer_checador_Tick(object sender, EventArgs e)
        {
            DateTime hoy = DateTime.Now;
            if (posponer == true)
            {
                /*Verificar si hay internet, si no posponerlo */
                if (AccesoInternet())
                {
                    _procesos("clientes", "", "", hoy.ToString("yyyy-MM-dd"));
                    posponer = false;
                    timer_checador.Interval = (1000 * 60); // cada  minuto
                    timer_checador.Start();
                }
                else
                {
                    posponer = true;
                    timer_checador.Interval = (1000 * 300); // cada  5 minuto
                    timer_checador.Start();
                }
            }
            else
            {
                if (hoy.Month + 1 < 13)
                { diaultimomes = new DateTime(hoy.Year, hoy.Month + 1, 1, hora, minuto, segundo).AddDays(-1); }
                else
                { diaultimomes = new DateTime(hoy.Year + 1, 1, 1, hora, minuto, segundo).AddDays(-1); }

                int result = DateTime.Compare(hoy, diaultimomes); // Compara si es antes, igual o despues de la hora (-1,0,1)
                var Hora = DateTime.Now.ToString("HH:mm");
                if ((hoy.ToString("dd-MMM-yyyy") == diaultimomes.ToString("dd-MMM-yyyy")) && (Hora == horainicial2))
                {                    
                    _procesos("clientes", "", "", hoy.ToString("yyyy-MM-dd"));
                }
            }

        }

        private void calibracionToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form2 View2 = new Form2("Correos calibracion");
            View2._gridview("correo_calibracion");
            View2.Show();
        }

        private void salidaToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            Form2 View2 = new Form2("Correos salida");
            View2._gridview("correo_salida");
            View2.Show();
        }

        private void facturaciónToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form2 View2 = new Form2("Correos facturación");
            View2._gridview("correo_facturacion");
            View2.Show();
        }

        private void cotizaciónToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form2 View2 = new Form2("Correos cotización");
            View2._gridview("correo_cotizacion");
            View2.Show();
        }

        private void logsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form3 View3 = new Form3();
            View3.Show();
        }

        private void reporteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form2 View2 = new Form2("Correos reportes");
            View2._gridview("correo_reporte");
            View2.Show();
        }

        private void metroButton10_Click(object sender, EventArgs e)
        {
            var m = new Random();
            int next = m.Next(0, 13);
            metroStyleManager1.Style = (MetroColorStyle)next;
        }

        private void metroButton13_Click(object sender, EventArgs e)
        {
            metroStyleManager1.Theme = metroStyleManager1.Theme == MetroThemeStyle.Light ? MetroThemeStyle.Dark : MetroThemeStyle.Light;
        }

        private void metroButtonenviar_Click(object sender, EventArgs e)
        {

            try
            {
                if (metroComboproceso.SelectedIndex != 0)
                {
                    if (metroComboproceso.Text == "clientes")
                    {

                        string fecha = metroDateTime1.Text;
                        DateTime date = Convert.ToDateTime(fecha);
                        string convert_date = date.ToString("yyyy-MM-dd");
                         _procesos(metroComboproceso.Text, metroComboproceso.Text.ToUpper(), "", convert_date);
                    }
                    else
                    {
                        _procesos(metroComboproceso.Text, metroComboproceso.Text.ToUpper(), metroTextBoxcorreo.Text);
                        metroTextBoxcorreo.Text = "";
                    }
                }
                MetroMessageBox.Show(this, "Se envio Exitosamente!", "Mensaje de notificación", MessageBoxButtons.OK, MessageBoxIcon.Question);
                metroComboproceso.SelectedIndex = 0;
            }
            catch (Exception)
            {
                MetroMessageBox.Show(this, "Error al enviar el correo, verificar datos. Si persite el error reportarlo al administrador!", "Mensaje de notificación", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }
        }

        private void metroButton1_Click(object sender, EventArgs e)
        {
            // Array [cliente,nogales,hermosillo,nogales,htmlbody]
            bool[] Arrcheck = { metroCheckBox1.Checked, metroCheckBox2.Checked, metroCheckBox3.Checked, metroCheckBox4.Checked, metroCheckBox6.Checked, metroCheckBox5.Checked };
            String[] Arrayquery =
            {
                "SELECT email FROM mypsa_bitacoramyp.view_usuarios where roles_id=10005 and activo='si';",
                "SELECT email FROM mypsa_bitacoramyp.view_usuarios where plantas_id=758  and activo='si';",
                "SELECT email FROM mypsa_bitacoramyp.view_usuarios where plantas_id=757  and activo='si';",
                "SELECT email FROM mypsa_bitacoramyp.view_usuarios where plantas_id=2341  and activo='si';",
                "",
                "",
            };
            string asunto = metroTextBox_asunto.Text;
            string body = metroTextBox_body.Text;

            for (int i = 0; i < Arrcheck.Length - 1; i++)
            {
                if (Arrcheck[i] == true)
                {
                    string contactos = "";
                    if (i == 4) // Opcion de otros
                    {
                        /*  Opcion  */
                        string correos = metroTextBox_correos.Text;
                        if (filecsv == true)
                        {
                            contactos = her_depurar_ejecutar._listaemails(dt2);
                        }
                        if (!String.IsNullOrEmpty(correos))
                        {
                            contactos += "," + correos;
                        }
                    }
                    else
                    {
                        DataTable dt_email = new DataTable();
                        dt_email = her_depurar_ejecutar._ejecutar(Arrayquery[i]);
                        contactos = her_depurar_ejecutar._listaemails(dt_email);
                    }

                    try
                    {
                        if (!String.IsNullOrEmpty(contactos)) { string retorno = correo._enviartemp(asunto, contactos, body, Arrcheck[5]); }
                        else { MetroMessageBox.Show(this, "Se ha detecto que no existen contactos, favor de verificar!", "Mensaje de notificación", MessageBoxButtons.OK, MessageBoxIcon.Warning); }

                        MetroMessageBox.Show(this, "Operacion exitosa!", "Mensaje de notificación", MessageBoxButtons.OK, MessageBoxIcon.Question);
                    }
                    catch (Exception)
                    {
                        MetroMessageBox.Show(this, "Se ha detecto un error al enviar mensaje, Contactarse con el Administrador!", "Mensaje de notificación", MessageBoxButtons.OK, MessageBoxIcon.Error);

                    }

                }
            }



        }

        private void metroLink1_Click(object sender, EventArgs e)
        {
            dt2.Clear();
            filecsv = false;
            metroLabelFile.Text = "Archivo limpio";
            metroTextBox_body.Text = "";
        }

        private void metroButton1_Click_1(object sender, EventArgs e)
        {
            try
            {
                using (OpenFileDialog ofd = new OpenFileDialog() { Filter = "CSV|*.csv", ValidateNames = true, Multiselect = false })
                {
                    if (ofd.ShowDialog() == DialogResult.OK)
                    {
                        dt2 = readCsv(ofd.FileName);
                        if (dt2.Rows.Count > 0)
                        {
                            metroLabelFile.Text = "Archivo listo, Contactos: # " + dt2.Rows.Count;
                            filecsv = true;
                        }
                    }
                }
            }
            catch (Exception)
            {
                MetroMessageBox.Show(this, "Error de operación!", "Mensaje de notificación", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public DataTable readCsv(string filename)
        {
            DataTable dt = new DataTable("Data");
            using (OleDbConnection cn = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\"" +
                    Path.GetDirectoryName(filename) + "\";Extended Properties='text; HDR=yes;FMT=Delimited(,)';"))
            {
                using (OleDbCommand cmd = new OleDbCommand(String.Format("select * from [{0}]", new FileInfo(filename).Name), cn))
                {
                    cn.Open();
                    using (OleDbDataAdapter adapter = new OleDbDataAdapter(cmd))
                    {
                        adapter.Fill(dt);
                    }
                }
            }
            return dt;
        }

        private void metroCheckBox6_CheckedChanged(object sender, EventArgs e)
        {
            if (metroCheckBox6.Checked == true)
            {
                metroTextBox_correos.Enabled = true;
                metroButtonFileCsv.Enabled = true;
            }
            else
            {
                metroTextBox_correos.Enabled = false;
                metroButtonFileCsv.Enabled = false;
            }
        }

        private void metroLink2_Click(object sender, EventArgs e)
        {
            try
            {
                using (OpenFileDialog ofd = new OpenFileDialog() { Filter = "HTML|*.html", ValidateNames = true, Multiselect = false })
                {
                    if (ofd.ShowDialog() == DialogResult.OK)
                    {
                        string bodyhtml = File.ReadAllText(ofd.FileName);
                        if (bodyhtml.Length > 0)
                        {
                            metroTextBox_body.Text = bodyhtml;
                        }
                    }
                }
            }
            catch (Exception)
            {
                MetroMessageBox.Show(this, "Error de operación!", "Mensaje de notificación", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            _setup();
        }

        private void Btn_ajustes_Click(object sender, EventArgs e)
        {
            Form4 View4 = new Form4();
            View4.Show();
        }

        private void metroComboproceso_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (metroComboproceso.Text == "cliente")
            {
                metroLabel11.Enabled = false;
                metroTextBoxcorreo.Enabled = false;
            }
            else
            {
                metroLabel11.Enabled = true;
                metroTextBoxcorreo.Enabled = true;
            }

        }

        private bool AccesoInternet()
        {
            try
            {
                System.Net.IPHostEntry host = System.Net.Dns.GetHostEntry("www.google.com");
                return true;
            }
            catch (Exception ex)
            {
                return false;
            }

        }

        private void _setup()
        {
            DateTime hoy = DateTime.Now;
            bool[] proceso = { false, false, false, false, false };// tecnico,calibracion,salida,factura,cliente

            /* Verificar la hora de envios programados */
            int result = DateTime.Compare(hoy, fechainicial); // Compara si es antes, igual o despues de la hora (-1,0,1)
            int result2 = DateTime.Compare(hoy, fechainicial.AddMinutes(20)); // Compara si es antes, igual o despues de la hora (-1,0,1)
            /* Si es antes de las 7:55 no hacer nada*/
            /* Si es despues de las 7:55 y 8:15 */
            if (result == 1 && result == 1)
            {
                /* Consultar si hay información de ese día */
                var query = "SELECT * FROM mypsa_recordatorios.registro where fecha like '%" + hoy.ToString("yyyy-MM-dd") + "%';";
                DataTable registro = new DataTable();
                registro = her_depurar_ejecutar._ejecutar(query, "2");
                /* Cuando no se encuentre registro se podra realizar el reenvio de recordatorios*/
                if (registro.Rows.Count > 0)
                {
                    /* Comprobando si falta enviar un recordatorio de los procesos o el cliente  */
                    foreach (DataRow row in registro.Rows)
                    {
                        var operacion = Convert.ToInt32(row["idoperacion"].ToString());
                        var estado = row["estado"].ToString();

                        if (estado == "exitoso")
                        {
                            if (operacion == 1)
                                proceso[0] = true;
                            if (operacion == 2)
                                proceso[1] = true;
                            if (operacion == 3)
                                proceso[2] = true;
                            if (operacion == 4)
                                proceso[3] = true;
                            if (operacion == 5)
                                proceso[4] = true;
                        }
                        // Si el valor es falso es porque hizo falta y se va ejecutar el envio de recordatorio
                    }
                }

                /* Una vez comprabado , Verificar si dia normal o es el ultimo día del mes para los recordatorios de los clientes*/
                if (hoy.Month + 1 < 13)
                { diaultimomes = new DateTime(hoy.Year, hoy.Month + 1, 1, hoy.Hour, hoy.Minute, hoy.Second).AddDays(-1); }
                else
                { diaultimomes = new DateTime(hoy.Year + 1, 1, 1, hoy.Hour, hoy.Minute, hoy.Second).AddDays(-1); }
                // es el ultimo día del mes y no se ha enviado
                if (hoy.Day == diaultimomes.Day && proceso[4] == false)
                {
                    /*Verificar si hay internet, si no posponerlo */
                    if (AccesoInternet())
                    {
                        _procesos("clientes", "", "", hoy.ToString("yyyy-MM-dd"));
                    }
                    else
                    {
                        posponer = true;
                        timer_checador.Interval = (1000 * 300); // cada  5 minuto
                        timer_checador.Start();
                    }
                }
                else
                {
                    /*Verificar si hay internet, si no posponerlo */
                    if (AccesoInternet())
                    {
                        if (hoy.DayOfWeek != DayOfWeek.Saturday && hoy.DayOfWeek != DayOfWeek.Sunday)
                        {
                            if (proceso[0] == false) //tecnico
                                _procesos("calib_tec");
                            if (proceso[1] == false) //calibracion
                                _procesos("calibracion", "Calibración");
                            if (proceso[2] == false) //salida
                                _procesos("salida", "Salida");
                            if (proceso[3] == false) //facturacion
                                _procesos("facturacion", "Facturación");

                        }
                    }
                    else
                    {
                        fechainicial = new DateTime(hoy.Year, hoy.Month, hoy.Day, hoy.Hour, hoy.Minute, 0).AddMinutes(5); // Parametro de 24 horas  
                    }
                }
            }
        }

    }
}
