using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ConfrontoListaFanta.ConfrontaFile
{
    class Confronter
    {
        private Excel.Workbook mioFileExcel { get; set; }

        private Excel.Workbook nuovoFileExcel { get; set; }

        public Confronter()
        {
        }


        public Excel.Worksheet ReadListone(string filePath, string sheetNme)
        {
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook WB = excelApp.Workbooks.Open(filePath);
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
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook WB = excelApp.Workbooks.Open(filePath);

            // statement get the workbookname  
            string ExcelWorkbookname = WB.Name;
            mioFileExcel = WB;
            // statement get the worksheet count  
            int worksheetcount = WB.Worksheets.Count;

            Excel.Worksheet wks = (Excel.Worksheet)WB.Worksheets[sheetNme];
            return wks;
        }
        public bool AggiungiMancanti(Excel.Worksheet Listone, Excel.Worksheet FileCustom)
        {
            try
            {
                //Prendo l'ultima riga nn nulla del foglio nuovo
                Excel.Range last = Listone.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
                int lastUsedRow = last.Row;

                //Prendo l'ultima riga nn nulla del mio foglio
                Excel.Range lastSalveMio = FileCustom.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
                int lastUsedRowMio = lastSalveMio.Row;

                //Rappresenta la prima riga nulla, ovvero quella dove evenutualmente inseriro i nuovi dati
                int RowToInsert = lastUsedRowMio + 2;


                for (int i = 2; i <= lastUsedRow; i++)
                {
                    bool exsist = false;
                    bool aggiornaSquadra = false;

                    //Rappresenta la riga che dovra essere non inserita, ma aggiornata con la nuova squadra in caso di cambio di maglia
                    int rigaDaAggiornare = 0;

                    //Leggo i valori dal file nuovo
                    var Name = ((Excel.Range)Listone.Cells[i, 4]).Value;
                    var Ruolo = ((Excel.Range)Listone.Cells[i, 2]).Value;
                    var Squadra = ((Excel.Range)Listone.Cells[i, 5]).Value;
                    var QuotazioneAttuale = ((Excel.Range)Listone.Cells[i, 6]).Value;

                    for (int j = 2; j <= lastUsedRowMio; j++)
                    {
                        //Leggo i vecchi valori dal mio file
                        string myName = ((Excel.Range)FileCustom.Cells[j, 2]).Value;
                        string myRuolo = ((Excel.Range)FileCustom.Cells[j, 1]).Value;
                        string mySquadra = ((Excel.Range)FileCustom.Cells[j, 3]).Value;

                        //Confronto
                        if (Name.Equals(myName))
                        {
                            exsist = true;
                            if (!Squadra.Equals(mySquadra))
                            {
                                aggiornaSquadra = true;
                                rigaDaAggiornare = j;
                            }
                            break;
                        }
                    }

                    if (!exsist)
                    {
                        //Devo Inserirlo
                        ((Excel.Range)FileCustom.Cells[RowToInsert, 1]).Value = Ruolo;
                        ((Excel.Range)FileCustom.Cells[RowToInsert, 2]).Value = Name;
                        ((Excel.Range)FileCustom.Cells[RowToInsert, 3]).Value = Squadra;
                        ((Excel.Range)FileCustom.Cells[RowToInsert, 4]).Value = QuotazioneAttuale;
                        ((Excel.Range)FileCustom.Rows[RowToInsert]).Interior.Color = XlRgbColor.rgbRed;
                        RowToInsert++;
                    }
                    else
                    {
                        if (aggiornaSquadra)
                        {
                            ((Excel.Range)FileCustom.Cells[rigaDaAggiornare, 3]).Value = Squadra;
                            ((Excel.Range)FileCustom.Cells[rigaDaAggiornare, 4]).Value = QuotazioneAttuale;
                            ((Excel.Range)FileCustom.Rows[rigaDaAggiornare]).Interior.Color = XlRgbColor.rgbYellowGreen;
                            rigaDaAggiornare = 0;
                        }
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
                mioFileExcel.Close();
                nuovoFileExcel.Close();
                return false;
            }
        }
        public bool RimuoviAndatiVia(Excel.Worksheet FileCustom, Excel.Worksheet Listone)
        {
            try
            {


                //Prendo l'ultima riga nn nulla del mio foglio
                Excel.Range last = FileCustom.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
                int lastUsedRowMio = last.Row;

                //Prendo l'ultima riga nn nulla del nuovo foglio
                Excel.Range lastSalveMio = Listone.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
                int lastUsedRow = lastSalveMio.Row;


                for (int i = 2; i <= lastUsedRowMio; i++)
                {
                    bool exsist = false;

                    //Leggo i valori dal mio vecchio file
                    var Name = ((Excel.Range)FileCustom.Cells[i, 2]).Value;
                    var Ruolo = ((Excel.Range)FileCustom.Cells[i, 1]).Value;
                    var Squadra = ((Excel.Range)FileCustom.Cells[i, 3]).Value;

                    if (Name == null)
                    {
                        break;
                    }
                    for (int j = 2; j <= lastUsedRow; j++)
                    {
                        //Leggo i vecchi valori dal nuovo file
                        string myName = ((Excel.Range)Listone.Cells[j, 4]).Value;
                        string myRuolo = ((Excel.Range)Listone.Cells[j, 2]).Value;
                        string mySquadra = ((Excel.Range)Listone.Cells[j, 5]).Value;

                        //Confronto
                        if (Name.Equals(myName))
                        {
                            exsist = true;
                            break;
                        }
                    }

                    //Se vado qui il giocatore non esiste nel nuovo foglio, lo rimuovo
                    if (!exsist)
                    {
                        //Devo Rimuovere dal mio foglio il giocatore (probabilmente è andato via perche esiste nel mio ma nn in quello aggiornato
                        ((Excel.Range)FileCustom.Rows[i]).Delete();
                    }
                }

                //Salvo
                mioFileExcel.Save();
                mioFileExcel.Close();
                nuovoFileExcel.Close();
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.ReadLine();
                mioFileExcel.Close();
                nuovoFileExcel.Close();
                return false;
            }
        }
    }
}
