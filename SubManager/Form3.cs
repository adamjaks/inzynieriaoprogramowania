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
    public partial class Form3 : Form
    {
        public Form3()
        {
            InitializeComponent();
        }
        Excel ex = new Excel();

        public void ReadCell(int i, int j)
        {
            ex.ReadCell(i, j);
        }
        public void WriteCell(int i, int j, string s)
        {
            Form2 tmp = new Form2();
            s = tmp.LabelText9;
        }
        
        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ex.SaveFile();
        }

        private void Form3_Load(object sender, EventArgs e)
        {
            label5.Text = ex.ReadCell(5, 2);
        }
    }
}
