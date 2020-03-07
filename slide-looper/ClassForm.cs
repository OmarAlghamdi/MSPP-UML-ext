using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace slide_looper
{
    public partial class ClassForm : Form
    {

        public String name { get; set; }
        public String[] fields { get; set; }
        public String[] ops { get; set; }

        public ClassForm()
        {
            InitializeComponent();
        }

        private void btn_cancel_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void btn_ok_Click(object sender, EventArgs e)
        {
            if (tb_name.Text == "")
            {
                label1.ForeColor = Color.Red;
                return;
            }
            if (tb_ops.Text == "")
            {
                label3.ForeColor = Color.Red;
                return;
            }
            name = tb_name.Text;
            fields = tb_fileds.Text.Split(';');
            ops = tb_ops.Text.Split(';');
            DialogResult = DialogResult.OK;
            Close();
        }
    }
}
