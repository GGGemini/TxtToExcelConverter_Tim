using System;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace TxtToExcelConverter_Tim
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            
            openFileDialog.Filter = "Text files(*.txt)|*.txt|All files(*.*)|*.*";

            if (openFileDialog.ShowDialog() == DialogResult.Cancel)
            {
                return;
            }

            string fileName = openFileDialog.FileName;

            richTextBox1.Text = File.ReadAllText(fileName, Encoding.GetEncoding(1251));
        }

        private void button2_Click(object sender, EventArgs e)
        {
            throw new System.NotImplementedException();
        }
    }
}
