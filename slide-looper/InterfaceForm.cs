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
    public partial class InterfaceForm : Form
    {
        public String name { get; set; }
        public String[] ops { get; set; }


        public InterfaceForm()
        {
            InitializeComponent();
        }
        
        private void btn_cancel_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void btn_ok_Click(object sender, EventArgs e)
        {
            name = tb_name.Text;
            ops = tb_ops.Text.Split(';');
            DialogResult = DialogResult.OK;
            Close();
        }
    }
}
