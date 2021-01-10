using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Reflection;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel.Application;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace SubManager
{
    public partial class Form2 : Form
    {
        public static string path1;
        Excel ex = new Excel(path1, 1);

        public Form2()
        {
            InitializeComponent();
        }

        
        public void OpenFile()
        {
            Excel excel = new Excel(path1, 1);           
        }
        public void ReadCell(int i, int j)
        {
            ex.ReadCell(i, j);
        }


        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void groupBox3_Enter(object sender, EventArgs e)
        {
                    
        }

        private void groupBox32_Enter(object sender, EventArgs e)
        {

        }

        private void groupBox2_Enter(object sender, EventArgs e)
        {
        }  

        private void Form2_Load(object sender, EventArgs e)
        {
            
                OpenFile();
                groupBox2.Text = ex.ReadCell(1, 22);
                groupBox3.Text = ex.ReadCell(3, 4);
                groupBox4.Text = ex.ReadCell(3, 5);
                groupBox5.Text = ex.ReadCell(3, 6);
                groupBox8.Text = ex.ReadCell(3, 7);
                groupBox7.Text = ex.ReadCell(3, 8);
                groupBox6.Text = ex.ReadCell(3, 9);
                groupBox9.Text = ex.ReadCell(3, 10);
                groupBox16.Text = ex.ReadCell(3, 11);
                groupBox15.Text = ex.ReadCell(3, 12);
                groupBox14.Text = ex.ReadCell(3, 13);
                groupBox13.Text = ex.ReadCell(3, 14);
                groupBox12.Text = ex.ReadCell(3, 15);
                groupBox11.Text = ex.ReadCell(3, 16);
                groupBox10.Text = ex.ReadCell(3, 17);
                groupBox23.Text = ex.ReadCell(3, 18);
                groupBox22.Text = ex.ReadCell(3, 19);
                groupBox21.Text = ex.ReadCell(3, 20);
                groupBox20.Text = ex.ReadCell(3, 21);
                groupBox19.Text = ex.ReadCell(3, 22);
                groupBox18.Text = ex.ReadCell(3, 23);
                groupBox17.Text = ex.ReadCell(3, 24);
                groupBox30.Text = ex.ReadCell(3, 25);
                groupBox29.Text = ex.ReadCell(3, 26);
                groupBox28.Text = ex.ReadCell(3, 27);
                groupBox27.Text = ex.ReadCell(3, 28);
                groupBox26.Text = ex.ReadCell(3, 29);
                groupBox25.Text = ex.ReadCell(3, 30);
                groupBox24.Text = ex.ReadCell(3, 31);
                groupBox33.Text = ex.ReadCell(3, 32);
                groupBox32.Text = ex.ReadCell(3, 33);
                groupBox31.Text = ex.ReadCell(3, 34);
                label1.Text = ex.ReadCell(1, 8);
                label21.Text = ex.ReadCell(4, 4);
                label16.Text = ex.ReadCell(4, 5);
                label17.Text = ex.ReadCell(4, 6);
                label20.Text = ex.ReadCell(4, 7);
                label19.Text = ex.ReadCell(4, 8);
                label18.Text = ex.ReadCell(4, 9);
                label22.Text = ex.ReadCell(4, 10);
                label29.Text = ex.ReadCell(4, 11);
                label28.Text = ex.ReadCell(4, 12);
                label27.Text = ex.ReadCell(4, 13);
                label26.Text = ex.ReadCell(4, 14);
                label25.Text = ex.ReadCell(4, 15);
                label24.Text = ex.ReadCell(4, 16);
                label23.Text = ex.ReadCell(4, 17);
                label36.Text = ex.ReadCell(4, 18);
                label35.Text = ex.ReadCell(4, 19);
                label34.Text = ex.ReadCell(4, 20);
                label33.Text = ex.ReadCell(4, 21);
                label32.Text = ex.ReadCell(4, 22);
                label31.Text = ex.ReadCell(4, 23);
                label30.Text = ex.ReadCell(4, 24);
                label43.Text = ex.ReadCell(4, 25);
                label42.Text = ex.ReadCell(4, 26);
                label41.Text = ex.ReadCell(4, 27);
                label40.Text = ex.ReadCell(4, 28);
                label39.Text = ex.ReadCell(4, 29);
                label38.Text = ex.ReadCell(4, 30);
                label37.Text = ex.ReadCell(4, 31);
                label46.Text = ex.ReadCell(4, 32);
                label45.Text = ex.ReadCell(4, 33);
                label44.Text = ex.ReadCell(4, 34);






        }

    }
}
