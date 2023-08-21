using Excel = Microsoft.Office.Interop.Excel;
using ExcelManager;
using Microsoft.Office.Interop.Excel;
using System;


namespace FantaGoat
{
    public class FantaGoatManager : ExcelModifier
    {
        private readonly int SlotIndex = 8;
        private readonly int PrezzoIndex = 13;

        public FantaGoatManager(string fileEliaPath, string filePathFantaGoat) : base(fileEliaPath, filePathFantaGoat)
        {
        }

        public override bool Allign(string sheetNameFileElia, string sheetNameFileToCopyFrom)
        {
            return AllineaPrezziESlot(GetSheet(sheetNameFileElia), ExcelToCopyFrom.GetSheet(sheetNameFileToCopyFrom));
        }

        private bool AllineaPrezziESlot(Worksheet SheetElia, Worksheet SheetFantaGoat)
        {
            try
            {
                //Prendo l'ultima riga nn nulla del foglio nuovo
                Excel.Range last = SheetFantaGoat.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing);
                int lastUsedRow = last.Row;

                //Prendo l'ultima riga nn nulla del mio foglio
                Excel.Range lastSalveMio = SheetElia.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing);
                int lastUsedRowMio = lastSalveMio.Row;

                //Rappresenta la prima riga nulla, ovvero quella dove evenutualmente inseriro i nuovi dati
                int RowToInsert = lastUsedRowMio + 2;


                for (int i = 2; i <= lastUsedRow; i++)
                {
                    bool exsist = false;

                    //Rappresenta la riga che dovra essere non inserita, ma aggiornata con la nuova squadra in caso di cambio di maglia
                    int rigaDaAggiornare = 0;

                    //Leggo i valori dal file nuovo
                    var Name = ((Excel.Range)SheetFantaGoat.Cells[i, 3]).Value as string ?? "dsa";
                    var Slot = ((Excel.Range)SheetFantaGoat.Cells[i, 1]).Value as string ?? "";
                    var FantaIndex = ((Excel.Range)SheetFantaGoat.Cells[i, 4]).Value;
                    var Prezzo = ((Excel.Range)SheetFantaGoat.Cells[i, 5]).Value;

                    for (int j = 2; j <= lastUsedRowMio; j++)
                    {
                        //Leggo i vecchi valori dal mio file
                        string myName = ((Excel.Range)SheetElia.Cells[j, 2]).Value as string ?? "asd";

                        //Confronto
                        if (!string.IsNullOrEmpty(Name) && !string.IsNullOrEmpty(myName) && Name.ToUpper().Contains(myName.Replace("'", "").ToUpper().Split()[0]))
                        {
                            exsist = true;
                            rigaDaAggiornare = j;

                            break;
                        }
                    }

                    if (exsist)
                    {
                        ((Excel.Range)SheetElia.Cells[rigaDaAggiornare, SlotIndex]).Value = Slot.Replace("° SLOT", "");
                        ((Excel.Range)SheetElia.Cells[rigaDaAggiornare, PrezzoIndex]).Value = Prezzo;
                        rigaDaAggiornare = 0;
                    }
                }

                //Salvo
                Save();
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return false;
            }
            finally
            {
                CloseBook();
                ExcelToCopyFrom.CloseBook();
                Quit();
                ExcelToCopyFrom.Quit();
            }
        }
    }
}
