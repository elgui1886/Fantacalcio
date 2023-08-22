using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelManager
{
    public class ExcelReader<T> where T : ExcelCell, new()
    {
    protected readonly Dictionary<string, Type> _typeAlias = new()
    {
        {  "bool" ,typeof(bool) },
        {  "double", typeof(double) },
        {  "string" , typeof(string) },
    };
        protected Application Application { get; set; }
        protected Workbook WorkBook { get; set; }
        protected Worksheet Sheet { get; set; }

        public ExcelReader(string filePath)
        {
            Application = new Application();
            WorkBook = Application.Workbooks.Open(filePath);
        }
        public Worksheet GetSheet(string sheetName)
        {
            Sheet = (Worksheet)WorkBook.Worksheets[sheetName];
            return Sheet;
        }

        public IEnumerable<T> SetColumsIndexesByNames(IEnumerable<T> columnsName)
        {
            foreach (var cell in columnsName)
            {
                cell.Index = FindColumnIndexByName(Sheet, cell.Name);

            }
            return columnsName;
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



        protected int FindColumnIndexByName(Worksheet worksheet, string columnName)
        {
            Excel.Range usedRange = worksheet.UsedRange;
            foreach (Excel.Range cell in usedRange.Rows[1].Cells)
            {
                if (cell.Value != null && cell.Value.ToString() == columnName)
                {
                    return cell.Column;
                }
            }
            return -1; // Colonna non trovata
        }

    }
}