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
using MetroFramework.Components;

namespace Myp_Email
{
    public partial class Form2 : MetroFramework.Forms.MetroForm
    {
        Class.Class_ejecutar ejecutar = new Class.Class_ejecutar();

        int id_contacto = 0;
        string table = "";       

        public Form2(string tipo = "")
        {
            InitializeComponent();            
            this.Text = tipo;
            MaximizeBox = false;
            _combobox();
            this.metroButton2.Enabled = false;
            this.metroButton3.Enabled = false;
            this.metroButton1.Enabled = true;
        }

        public void _gridview(string tipo = "")
        {
            table = tipo;
            DataTable dt = new DataTable();
            DataTable dt_copy = new DataTable(); //Copia el contenido, porque cuando combobox entra a pedir otra vez al metodo se genera otro datatable. y el gridview esta alimentado con el datatable directamente
            dt = ejecutar._ejecutar("select id,nombre as Nombre,email,sucursal as Planta,id_sucursal from view_" + table + ";", "2");
            dt_copy = dt.Copy();
            //tabla.DataSource = dt_copy;
            metroGridmodulo.DataSource = dt_copy;
            metroGridmodulo.Columns[0].Visible = false;
            metroGridmodulo.Columns[1].Width = 110;
            metroGridmodulo.Columns[2].Width = 150;
            metroGridmodulo.Columns[3].Width = 80;
            metroGridmodulo.Columns[4].Visible = false;
        }

        public void _combobox()
        {
            DataTable dt = new DataTable();
            DataTable dt_copy = new DataTable();
            dt = ejecutar._ejecutar("select id,nombre from sucursal;", "2");
            dt_copy = dt.Copy();
            metroComboBox1.DataSource = dt_copy;
            metroComboBox1.DisplayMember = "nombre";
            metroComboBox1.ValueMember = "id";
        }        

        private Boolean _validar()
        {
            var result = true;
            if (String.IsNullOrEmpty(this.metroTextBox1.Text) == true && String.IsNullOrEmpty(this.metroTextBox2.Text) == true)
            {
                result = false;
                this.metroTextBox1.Focus();
            }
            return result;
        }

        private int duplicado(string query = "")
        {
            DataTable dt = new DataTable();
            dt = ejecutar._ejecutar(query, "2");
            int count = int.Parse(dt.Rows[0]["count"].ToString());
            return count;

        }

        private void metroButton1_Click(object sender, EventArgs e)
        {
            string tipo = table;
            if (_validar() == true)
            {
                string values = String.Format("'{0}','{1}','{2}'", this.metroTextBox1.Text, this.metroTextBox2.Text, this.metroComboBox1.SelectedValue);
                try
                {

                    if (duplicado("select count(correo) as count from " + tipo + " where correo='" + this.metroTextBox2.Text + "';") == 0)
                    {
                        ejecutar._add(tipo, values);                        
                        MetroMessageBox.Show(this, "Se agrego Exitosamente!", "Mensaje de notificación", MessageBoxButtons.OK, MessageBoxIcon.Question);
                        _gridview(tipo);
                        this.metroTextBox1.Text = "";
                        this.metroTextBox2.Text = "";
                        this.metroButton2.Enabled = false;
                        this.metroButton3.Enabled = false;
                    }
                    else
                    {
                        this.metroTextBox2.Focus();
                        MetroMessageBox.Show(this, "Correo duplicado!", "Mensaje de notificación", MessageBoxButtons.OK, MessageBoxIcon.Warning);                      
                    }
                }
                catch (Exception)
                {
                    MetroMessageBox.Show(this, "Error de operación!", "Mensaje de notificación", MessageBoxButtons.OK, MessageBoxIcon.Error);                    
                }
            }
            else
            {
                MetroMessageBox.Show(this, "Error de operación!", "Mensaje de notificación", MessageBoxButtons.OK, MessageBoxIcon.Error);                               
            }
        }

        private void metroButton2_Click(object sender, EventArgs e)
        {
            string tipo = table;
            if (_validar() == true)
            {
                string values = String.Format(" nombre='{0}', correo='{1}', id_sucursal='{2}' where id='{3}'", this.metroTextBox1.Text, this.metroTextBox2.Text, this.metroComboBox1.SelectedValue, id_contacto);
                try
                {
                    if (duplicado("select count(correo) as count from " + tipo + " where correo='" + this.metroTextBox2.Text + "' and id !=" + id_contacto + ";") == 0)
                    {
                        ejecutar._update(tipo, values);
                        MetroMessageBox.Show(this, "Se edito Exitosamente!", "Mensaje de notificación", MessageBoxButtons.OK, MessageBoxIcon.Question);                        
                        _gridview(tipo);
                        this.metroTextBox1.Text = "";
                        this.metroTextBox2.Text = "";
                        this.metroButton2.Enabled = false;
                        this.metroButton3.Enabled = false;
                    }
                    else
                    {
                        this.metroTextBox2.Focus();
                        MetroMessageBox.Show(this, "Correo duplicado!", "Mensaje de notificación", MessageBoxButtons.OK, MessageBoxIcon.Warning);                        
                    }
                }
                catch (Exception)
                {
                    MetroMessageBox.Show(this, "Error de operación!", "Mensaje de notificación", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                MetroMessageBox.Show(this, "Error de operación!", "Mensaje de notificación", MessageBoxButtons.OK, MessageBoxIcon.Error);                
            }
        }

        private void metroButton3_Click(object sender, EventArgs e)
        {
            string tipo = table;
            string values = String.Format("id='{0}'", id_contacto);
            try
            {
                ejecutar._delete(tipo, values);                
                MetroMessageBox.Show(this, "Se elimino Exitosamente!", "Mensaje de notificación", MessageBoxButtons.OK, MessageBoxIcon.Question);                
                _gridview(tipo);
                this.metroTextBox1.Text = "";
                this.metroTextBox2.Text = "";
                this.metroButton2.Enabled = false;
                this.metroButton3.Enabled = false;
                this.metroButton1.Enabled = true;
                this.metroTextBox1.Focus();
            }
            catch (Exception)
            {
                MetroMessageBox.Show(this, "Error de operación!", "Mensaje de notificación", MessageBoxButtons.OK, MessageBoxIcon.Error);                
            }
        }

        private void metroGridmodulo_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (metroGridmodulo.CurrentRow.Index != -1)
            {
                id_contacto = int.Parse(metroGridmodulo.CurrentRow.Cells[0].Value.ToString());
                this.metroTextBox1.Text = metroGridmodulo.CurrentRow.Cells[1].Value.ToString();
                this.metroTextBox2.Text = metroGridmodulo.CurrentRow.Cells[2].Value.ToString();

                this.metroComboBox1.SelectedIndex = int.Parse(metroGridmodulo.CurrentRow.Cells[4].Value.ToString()) - 1;

                this.metroButton2.Enabled = true;
                this.metroButton3.Enabled = true;
            }
        }
    }
}
