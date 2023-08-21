using Excel = Microsoft.Office.Interop.Excel;
using ExcelManager;
using Microsoft.Office.Interop.Excel;
using System;

namespace CostoAstaPrecedente
{
    //*
    // Questo manager si occupa di allineare i prezzi della ultima asta
    // Attenzione!! Il tool si aspetta determinate celle in deterinate posizioni. Se tali posizioni cambiano, il tool va aggiornato
    // Per riferimenti sulla posizione delle celle confrontare con file anni passati e/o verificare che il nome celle combaci di anno in anno
    //*
    public class CostoAstaPrecedenteManager : ExcelModifier
    {
        private readonly int QAPIndex = 17;
        public CostoAstaPrecedenteManager(string mioFilePath, string fileRoseAstaPrecedente) : base(mioFilePath, fileRoseAstaPrecedente)
        {
        }

        public override bool Allign(string sheetNameFileElia, string sheetNameFileToCopyFrom)
        {
            return AllineaCostoAnnoPrecedente(GetSheet(sheetNameFileElia), ExcelToCopyFrom.GetSheet(sheetNameFileToCopyFrom));
        }

        private bool AllineaCostoAnnoPrecedente(Worksheet SheetElia, Worksheet SheetAstaPrecedente)
        {
            try
            {
                //Prendo l'ultima riga nn nulla del foglio nuovo
                Excel.Range last = SheetAstaPrecedente.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing);
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
                    var Name = ((Excel.Range)SheetAstaPrecedente.Cells[i, 2]).Value as string ?? "dsa";
                    var Prezzo = ((Excel.Range)SheetAstaPrecedente.Cells[i, 4]).Value;

                    for (int j = 2; j <= lastUsedRowMio; j++)
                    {
                        //Leggo i vecchi valori dal mio file
                        string myName = ((Excel.Range)SheetElia.Cells[j, 2]).Value as string ?? "asd";

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
                        ((Excel.Range)SheetElia.Cells[rigaDaAggiornare, QAPIndex]).Value = Prezzo;
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
