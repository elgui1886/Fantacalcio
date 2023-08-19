using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ConfrontoListaFanta.ConfrontaFile
{
    public class AllineaPrezzi
    {
        private Excel.Workbook mioFileExcel { get; set; }

        private Excel.Workbook nuovoFileExcel { get; set; }

        private Excel.Application excelApp1 { get; set; }

        private Excel.Application excelApp2 { get; set; }

        public AllineaPrezzi()
        {
        }


        public Excel.Worksheet ReadFantaculo(string filePath, string sheetNme)
        {
            excelApp1 = new Excel.Application();
            Excel.Workbook WB = excelApp1.Workbooks.Open(filePath);
            nuovoFileExcel = WB;
            // statement get the workbookname  
            string ExcelWorkbookname = WB.Name;

            // statement get the worksheet count  
            int worksheetcount = WB.Worksheets.Count;

            Excel.Worksheet wks = (Excel.Worksheet)WB.Worksheets[sheetNme];
            return wks;
        }
        public Excel.Worksheet ReadFileExcelMio(string filePath, string sheetNme)
        {
            excelApp2 = new Excel.Application();
            Excel.Workbook WB = excelApp2.Workbooks.Open(filePath);

            // statement get the workbookname  
            string ExcelWorkbookname = WB.Name;
            mioFileExcel = WB;
            // statement get the worksheet count  
            int worksheetcount = WB.Worksheets.Count;

            Excel.Worksheet wks = (Excel.Worksheet)WB.Worksheets[sheetNme];
            return wks;
        }

        public bool AllineaPrezziESlot(Excel.Worksheet Fantaculo, Excel.Worksheet Mio)
        {
            try
            {
                //Prendo l'ultima riga nn nulla del foglio nuovo
                Excel.Range last = Fantaculo.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
                int lastUsedRow = last.Row;

                //Prendo l'ultima riga nn nulla del mio foglio
                Excel.Range lastSalveMio = Mio.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
                int lastUsedRowMio = lastSalveMio.Row;

                //Rappresenta la prima riga nulla, ovvero quella dove evenutualmente inseriro i nuovi dati
                int RowToInsert = lastUsedRowMio + 2;


                for (int i = 2; i <= lastUsedRow; i++)
                {
                    bool exsist = false;

                    //Rappresenta la riga che dovra essere non inserita, ma aggiornata con la nuova squadra in caso di cambio di maglia
                    int rigaDaAggiornare = 0;

                    //Leggo i valori dal file nuovo
                    var Name = (((Excel.Range)Fantaculo.Cells[i, 1]).Value as string);
                    var Slot = ((Excel.Range)Fantaculo.Cells[i, 6]).Value;
                    var PrezzoFC = ((Excel.Range)Fantaculo.Cells[i, 5]).Value;
                    var PrezzoAsta = ((Excel.Range)Fantaculo.Cells[i, 4]).Value;

                    for (int j = 2; j <= lastUsedRowMio; j++)
                    {
                        //Leggo i vecchi valori dal mio file
                        string myName = ((Excel.Range)Mio.Cells[j, 2]).Value;

                        //Confronto
                        if (!string.IsNullOrEmpty(Name) && !string.IsNullOrEmpty(myName) && Name.ToUpper().Equals(myName.ToUpper()))
                        {
                            exsist = true;
                            rigaDaAggiornare = j;
                            
                            break;
                        }
                    }

                    if (exsist)
                    {
                        ((Excel.Range)Mio.Cells[rigaDaAggiornare, 6]).Value = Slot;
                        ((Excel.Range)Mio.Cells[rigaDaAggiornare, 7]).Value = PrezzoFC;
                        ((Excel.Range)Mio.Cells[rigaDaAggiornare, 8]).Value = PrezzoAsta;
                        rigaDaAggiornare = 0;
                    }              
                }

                //Salvo
                mioFileExcel.Save();
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.ReadLine();
                return false;
            }
            finally
            {
                mioFileExcel.Close();
                nuovoFileExcel.Close();
                excelApp1.Quit();
                excelApp2.Quit();
            }
        }
    }
}
