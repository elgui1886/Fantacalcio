using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelManager
{
    public class ExcelReader<T> where T : ExcelCell, new()
    {
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
            Sheet =(Worksheet)WorkBook.Worksheets[sheetName];
            return Sheet;
        }

        public virtual List<T> GetColumsIndexesByNames(string[] columnName)
        {
            var mapper = new List<T>();   
            foreach (var name in columnName)
            {
                var index = FindColumnIndexByName(Sheet, name);
                if (index != -1)
                {
                    mapper.Add(new T { Name = name, Index = index });
                }
            }       
            return mapper;
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