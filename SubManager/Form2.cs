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
    public partial class Form2 : Form
    {
        public static string path1;
        Excel ex = new Excel(path1, 1);
        public Form2()
        {
            InitializeComponent();
        }


        public void ReadCell(int i, int j)
        {
            ex.ReadCell(i, j);
        }
        public string LabelText9
        {
            get
            {
                return this.label9.Text;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void groupBox3_Enter(object sender, EventArgs e)
        {
                    
        }

        private void Form2_Load(object sender, EventArgs e)
        {
               
                groupBox31.Hide();
                groupBox32.Hide();
                groupBox33.Hide();

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

            if (ex.ReadCell(3, 34) != null)
            {
                groupBox31.Show();
                groupBox31.Text = ex.ReadCell(3, 34);
                string openHours = ex.ReadCell(5, 34).ToString();
                string[] openHoursTab = openHours.Split(' ', '-');
                if (label9.Text.Equals(openHoursTab[0]) && label10.Text.Equals(openHoursTab[3]))
                    label44.ForeColor = Color.Green;
                else
                    label44.ForeColor = Color.Red;

                
            }
            if (ex.ReadCell(3, 33) != null)
            {
                groupBox32.Show();
                groupBox32.Text = ex.ReadCell(3, 33);
                string openHours = ex.ReadCell(5, 33).ToString();
                string[] openHoursTab = openHours.Split(' ', '-');
                if (label9.Text.Equals(openHoursTab[0]) && label10.Text.Equals(openHoursTab[3]))
                    label45.ForeColor = Color.Green;
                else
                    label45.ForeColor = Color.Red;
            }
            if (ex.ReadCell(3, 32) != null)
            {
                groupBox33.Show();
                groupBox33.Text = ex.ReadCell(3, 32);
                string openHours = ex.ReadCell(5, 32).ToString();
                string[] openHoursTab = openHours.Split(' ', '-');
                if (label9.Text.Equals(openHoursTab[0]) && label10.Text.Equals(openHoursTab[3]))
                    label46.ForeColor = Color.Green;
                else
                    label46.ForeColor = Color.Red;
            }



                label47.Text = ex.ReadCell(1, 8);
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

        private void button4_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void label21_Click(object sender, EventArgs e)
        {
            Form3 f3 = new Form3();
            f3.Show();

        }

        private void label35_Click(object sender, EventArgs e)
        {
            Form3 f3 = new Form3();
            f3.Show();
        }

        private void label16_Click(object sender, EventArgs e)
        {
            Form3 f3 = new Form3();
            f3.Show();
        }

        private void label17_Click(object sender, EventArgs e)
        {
            Form3 f3 = new Form3();
            f3.Show();
        }

        private void label20_Click(object sender, EventArgs e)
        {
            Form3 f3 = new Form3();
            f3.Show();
        }

        private void label19_Click(object sender, EventArgs e)
        {
            Form3 f3 = new Form3();
            f3.Show();
        }

        private void label18_Click(object sender, EventArgs e)
        {
            Form3 f3 = new Form3();
            f3.Show();
        }

        private void label22_Click(object sender, EventArgs e)
        {
            Form3 f3 = new Form3();
            f3.Show();
        }

        private void label29_Click(object sender, EventArgs e)
        {
            Form3 f3 = new Form3();
            f3.Show();
        }

        private void label28_Click(object sender, EventArgs e)
        {
            Form3 f3 = new Form3();
            f3.Show();
        }

        private void label27_Click(object sender, EventArgs e)
        {
            Form3 f3 = new Form3();
            f3.Show();
        }

        private void label26_Click(object sender, EventArgs e)
        {
            Form3 f3 = new Form3();
            f3.Show();
        }

        private void label25_Click(object sender, EventArgs e)
        {
            Form3 f3 = new Form3();
            f3.Show();
        }

        private void label24_Click(object sender, EventArgs e)
        {
            Form3 f3 = new Form3();
            f3.Show();
        }

        private void label23_Click(object sender, EventArgs e)
        {
            Form3 f3 = new Form3();
            f3.Show();
        }

        private void label36_Click(object sender, EventArgs e)
        {
            Form f3 = new Form3();
            f3.Show();
        }

        private void label34_Click(object sender, EventArgs e)
        {
            Form3 f3 = new Form3();
            f3.Show();
        }

        private void label33_Click(object sender, EventArgs e)
        {
            Form3 f3 = new Form3();
            f3.Show();
        }

        private void label32_Click(object sender, EventArgs e)
        {
            Form3 f3 = new Form3();
            f3.Show();
        }

        private void label31_Click(object sender, EventArgs e)
        {
            Form3 f3 = new Form3();
            f3.Show();
        }

        private void label30_Click(object sender, EventArgs e)
        {
            Form3 f3 = new Form3();
            f3.Show();
        }

        private void label43_Click(object sender, EventArgs e)
        {
            Form3 f3 = new Form3();
            f3.Show();
        }

        private void label42_Click(object sender, EventArgs e)
        {
            Form3 f3 = new Form3();
            f3.Show();
        }

        private void label41_Click(object sender, EventArgs e)
        {
            Form3 f3 = new Form3();
            f3.Show();
        }

        private void label40_Click(object sender, EventArgs e)
        {
            Form3 f3 = new Form3();
            f3.Show();
        }

        private void label39_Click(object sender, EventArgs e)
        {
            Form3 f3 = new Form3();
            f3.Show();
        }

        private void label38_Click(object sender, EventArgs e)
        {
            Form3 f3 = new Form3();
            f3.Show();
        }

        private void label37_Click(object sender, EventArgs e)
        {
            Form3 f3 = new Form3();
            f3.Show();
        }

        private void label46_Click(object sender, EventArgs e)
        {
            Form3 f3 = new Form3();
            f3.Show();
        }

        private void label45_Click(object sender, EventArgs e)
        {
            Form3 f3 = new Form3();
            f3.Show();
        }

        private void label44_Click(object sender, EventArgs e)
        {
            Form3 f3 = new Form3();
            f3.Show();
        }
    }
}
