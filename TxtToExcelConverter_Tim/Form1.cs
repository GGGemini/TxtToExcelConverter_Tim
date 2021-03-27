using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Windows.Forms;
using ClosedXML.Excel;
using TxtToExcelConverter_Tim.Logic;
using TxtToExcelConverter_Tim.Models;

namespace TxtToExcelConverter_Tim
{
    public partial class Form1 : Form
    {
        private string fileName;
        
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

            fileName = openFileDialog.FileName;

            richTextBox1.Text = File.ReadAllText(fileName, Encoding.GetEncoding(1251));

            fileName = fileName
                        .Remove(fileName.LastIndexOf('.'))
                        .Substring(fileName.LastIndexOf('\\') + 1);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(fileName))
                return;

            SaveFileDialog sF = new SaveFileDialog()
            {
                Filter = "Excel files(*.xlsx)|*.xlsx|All files(*.*)|*.*",
                FileName = $"{fileName}.xlsx"
            };
            
            if (sF.ShowDialog() == DialogResult.OK)
            {
                Stream stream = null;

                // получение моделей
                TableModel[] models = TextLogic.GetModelsFromString(richTextBox1.Text);

                // создание excel
                XLWorkbook wb = ExcelLogic.GetExcel(models);

                // сохранение
                if (File.Exists(sF.FileName))
                {
                    // искал решение красивее, но не нашёл, как узнать флаг доступен для записи или нет
                    try
                    {
                        // выясняем, не открыт ли файл у пользователя
                        using (FileStream fs = File.OpenWrite(sF.FileName))
                        {

                        }
                    }
                    catch (IOException)
                    {
                        // если открыт, то сохраняем под другим именем
                        sF.FileName = sF.FileName.Replace(fileName, DateTime.Now.ToLocalTime().ToString("dd-MM-yyyy_HH-mm-ss-ff"));
                    }
                }

                stream = sF.OpenFile();

                wb.SaveAs(stream);
                stream.Close();
            }
        }
    }
}
