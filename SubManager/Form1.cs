using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SubManager
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            labelUser.Text = Environment.UserName + " (" + Environment.MachineName + ")";
        }

        private void button3_Click(object sender, EventArgs e)
        {
            OpenFileDialog importFileDialog = new OpenFileDialog();
            importFileDialog.Filter = "*.xls|*.xlsx";
            importFileDialog.Multiselect = false;

            if (importFileDialog.ShowDialog() == DialogResult.OK)
            {
                button3.Text = importFileDialog.FileName;
            }
        }
    }
}
