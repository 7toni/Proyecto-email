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
    public partial class Form3 : MetroFramework.Forms.MetroForm
    {
        Class.Class_ejecutar ejecutar = new Class.Class_ejecutar();
        DataTable dt = new DataTable();
        DataTable dt_copy = new DataTable(); //Copia el contenido, porque cuando combobox entra a pedir otra vez al metodo se genera otro datatable. y el gridview esta alimentado con el datatable directamente

        public Form3()
        {
            InitializeComponent();
            MaximizeBox = false;
            this.MaximizeBox = false;
            metroComboBox1.SelectedIndex = 0;
            string limit = metroComboBox1.SelectedItem.ToString();
            _gridview(limit);
        }

        public void _gridview(string limit)
        {
            dt = ejecutar._ejecutar("select * from view_registro order by id desc limit " + limit + ";", "2");
            dt_copy = dt.Copy();
            metroGridlogs.DataSource = dt_copy;
            metroGridlogs.Columns[0].Visible = false;
            metroGridlogs.Columns[1].Visible = false;
            metroGridlogs.Columns[3].Visible = false;
            metroGridlogs.Columns[7].Visible = false;
            metroGridlogs.Columns[6].Width = 125;
        }

        private void metroGridlogs_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {

                var contenido = "*************** DETALLES *************** - " + this.metroGridlogs.CurrentRow.Cells[7].Value.ToString();
                var concatenar = "";
                for (int i = 0; i < contenido.Length; i++)
                {
                    if (contenido[i].ToString() == "-")
                    {
                        concatenar += "\r\n";
                    }
                    else
                    {
                        concatenar += contenido[i].ToString();
                    }
                }

                metroDetalle.Text = concatenar;
            }
        }

        private void metroComboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            string limit = metroComboBox1.SelectedItem.ToString();
            _gridview(limit);

        }

    }
}
