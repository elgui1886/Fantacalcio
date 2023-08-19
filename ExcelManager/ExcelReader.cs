using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelManager
{
    public class ExcelReader
    {
        private Excel.Application Application { get; set; }
        private Excel.Workbook WorkBook { get; set; }

        public ExcelReader(string filePath)
        {
            Application = new Excel.Application();
            WorkBook = Application.Workbooks.Open(filePath);
        }
        public Excel.Worksheet GetSheet(string sheetNme)
        {          
            return (Excel.Worksheet)WorkBook.Worksheets[sheetNme];
        }

        public void Save()
        {
            WorkBook.Save();
        }

        public void CloseBook()
        {
            WorkBook.Close();
        }

        public void Quit()
        {
            Application.Quit();
        }

    }
}