using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;

namespace gyak12
{
    public partial class Form1 : Form
    {
        Excel.Application xlApp;
        Excel.Workbook xlWb;
        Excel.Worksheet xlSheet;
        public Form1()
        {
            InitializeComponent();

            try
            {
                xlApp = new Excel.Application();

                xlWb = xlApp.Workbooks.Add(Missing.Value);

                xlSheet = xlWb.ActiveSheet;

                //CreateTable();

                xlApp.Visible= true;
                xlApp.UserControl = true;
            }
            catch (Exception ex)
            {
                string errMsg = string.Format("Error: {0}\nLine: {1}", ex.Message, ex.Source);
                MessageBox.Show(errMsg, "Error");

                xlWb.Close(false, Type.Missing, Type.Missing);
                xlApp.Quit();
                xlWb = null;
                xlApp = null;
            }
        }

        void CreateTable()
        {
            string[] fejlecek = new string[]
            {
                "Kérdés",
                "1. válasz",
                "2. válasz",
                "3. válasz",
                "Helyes válasz",
                "kép"
            };

            for (int i = 0; i < fejlecek.Length; i++)
            {
                xlSheet.Cells[1, 1] = fejlecek[0];
            }
        }

    }

}