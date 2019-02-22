using MetroFramework;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Myp_Email
{
    public partial class Form4 : MetroFramework.Forms.MetroForm
    {
        Class.Class_ejecutar ejecutar = new Class.Class_ejecutar();
        Class.Class_email email = new Class.Class_email();

        public Form4()
        {
            InitializeComponent();
            MaximizeBox = false;
            check_ssl.Checked = false;
            consultainfo();
        }

        private void consultainfo()
        {
            DataTable dt = new DataTable();
            DataTable dt_copy = new DataTable();
            dt = ejecutar._ejecutar("select * from view_correo_envio", "2");

            txtb_nombre.Text = dt.Rows[0]["nombre"].ToString();
            txtb_usuario.Text = dt.Rows[0]["usuario"].ToString();
            txtb_password.Text = dt.Rows[0]["password"].ToString();
            txtb_servidor.Text = dt.Rows[0]["servidor"].ToString();
            txtb_puerto.Text = dt.Rows[0]["puerto"].ToString();
            if (dt.Rows[0]["requiere_ssl"].ToString() == "1")
            {
                check_ssl.Checked = true;
            }
            else
            {
                check_ssl.Checked = false;
            }
            

        }

        private void metroButton2_Click(object sender, EventArgs e)
        {
            string tipo = "correo_envio";
            int requiere_ssl = 0;
            if (check_ssl.Checked == true)
            {
                requiere_ssl = 1;
            }            
            string values = String.Format(" nombre='{0}', usuario='{1}', password='{2}', servidor='{3}', puerto={4}, requiere_ssl={5} where id=1;", this.txtb_nombre.Text, this.txtb_usuario.Text, this.txtb_password.Text, this.txtb_servidor.Text, this.txtb_puerto.Text, requiere_ssl);
            try
            {
                string retorno = ejecutar._update(tipo, values);
                if (retorno == "Exitoso")
                {
                    MetroMessageBox.Show(this, "Se edito Exitosamente!", "Mensaje de notificación", MessageBoxButtons.OK, MessageBoxIcon.Question);
                    consultainfo();
                }
            }
            catch (Exception)
            {
                MetroMessageBox.Show(this, "Error de operación!", "Mensaje de notificación", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
           
            
        }

        private void metroButton1_Click(object sender, EventArgs e)
        {
            string toadd = "";
            string bccadd = "";            
            bool html = true;
            string bodyhtml = "Mensaje de correo electrónico enviado automáticamente por MyPSA Recordatorios para comprobar la configuración de su cuenta. ";
            string asunto = "Mypsa <Mensaje de prueba>";

            string retorno = email._enviodecorreo(asunto,toadd,bccadd,html,bodyhtml);
            if (retorno == "Exitoso")
            {
                MetroMessageBox.Show(this, "Se Envío Exitosamente!. Favor de revisar su bandeja de entrada. Gracias.", "Mensaje de notificación", MessageBoxButtons.OK, MessageBoxIcon.Question);
            }
            else {
                MetroMessageBox.Show(this, "Error de operación!. Favor de revisar las configuraciones. Gracias.", "Mensaje de notificación", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
