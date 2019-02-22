using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Myp_Email
{
    public partial class frmMain : MetroFramework.Forms.MetroForm
    {
        int i = 0;
        public frmMain()
        {
            InitializeComponent();
            timer1.Start();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            metroLabel1.Text = string.Format("Cargando {0} %", i++);
            this.metroProgressBar1.Value++;
            this.metroProgressBar1.Update();
            if (i == 101)
            {
                timer1.Stop();
            }
        }
    }
}
