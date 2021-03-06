using System;
using System.Text;
using System.Windows.Forms;

namespace TxtToExcelConverter_Tim
{
    static class Program
    {
        /// <summary>
        ///  The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            // чтобы читать кодировку с файла
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            
            Application.SetHighDpiMode(HighDpiMode.SystemAware);
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
        }
    }
}
