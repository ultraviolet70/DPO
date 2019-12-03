using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DPO
{
    public partial class FormScreen : Form
    {
        public FormScreen()
        {
            InitializeComponent();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            try
            {
                progressBar1.Value = progressBar1.Value + 10;
            }
            catch { }
        }

        private void timer2_Tick(object sender, EventArgs e)
        {
            FormMain formMain = new FormMain();
            formMain.Show();
            this.Hide();
            timer2.Enabled = false;
        }
    }
}
