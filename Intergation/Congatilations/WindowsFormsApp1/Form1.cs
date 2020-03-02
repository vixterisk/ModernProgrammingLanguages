using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void folderBrowserDialog1_HelpRequest(object sender, EventArgs e)
        {

        }

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }
        
        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            label1.Text = "";
            label2.Text = "";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                excelFileName.Text = openFileDialog1.FileName;
            }
        }

        private void excelFileName_TextChanged(object sender, EventArgs e)
        {
            if (excelFileName.Text != "")
            {
                label1.Text = Generator.CreateCongratilation(excelFileName.Text);
                if (!Generator.UniquePossibility) label2.Text = "Пополните список пожеланий/групп для уникальности поздравлений.";
            }
        }
    }
}
