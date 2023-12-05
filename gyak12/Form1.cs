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

                CreateTable();

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
                xlSheet.Cells[1, i+1] = fejlecek[i];
            }

            Models.HajosContext context = new Models.HajosContext();
            var mindenkerdes = context.Questions.ToList();

            object[,] adattomb = new object[mindenkerdes.Count(), fejlecek.Count()];

            for (int i = 0; i < mindenkerdes.Count(); i++)
            {
                adattomb[i, 0] = mindenkerdes[i].Question1;
                adattomb[i, 1] = mindenkerdes[i].Answer1;
                adattomb[i, 2] = mindenkerdes[i].Answer2;
                adattomb[i, 3] = mindenkerdes[i].Answer3;
                adattomb[i, 4] = mindenkerdes[i].CorrectAnswer;
                adattomb[i, 5] = mindenkerdes[i].Image;
            }

            int sorokSzáma = adattomb.GetLength(0);
            int oszlopokSzáma = adattomb.GetLength(1);

            Excel.Range adatRange = xlSheet.get_Range("A2", Type.Missing).get_Resize(sorokSzáma, oszlopokSzáma);
            adatRange.Value2 = adattomb;
            adatRange.Columns.AutoFit();
        }

    }

}