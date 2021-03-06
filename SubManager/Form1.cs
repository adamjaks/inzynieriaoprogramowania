﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

namespace SubManager
{
    public partial class Form1 : Form
    {
        string logData;

        public Form1()
        {
            InitializeComponent();
            logData = Environment.UserName + " (" + Environment.MachineName + ")";
            labelUser.Text = logData;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Title = " Otwórz plik ";
            openFileDialog1.Filter = "Excel Files|*.xlsx";
            openFileDialog1.ShowDialog();
            Form2 f2 = new Form2();
            f2.Show();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            string path = "../../../SM_logs.txt";
            StreamWriter sw = File.AppendText(path);
            sw.WriteLine(logData + " (" + DateTime.Now + ")");
            sw.Close();
        }
    }
}
 