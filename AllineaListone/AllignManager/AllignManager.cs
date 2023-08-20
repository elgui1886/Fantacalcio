using Excel = Microsoft.Office.Interop.Excel;
using ExcelManager;
using Microsoft.Office.Interop.Excel;
using System;

namespace AllineaListone
{
    //*
    // Questo manager si occupa di allineare il listone di Elia con quello di Fantagazzetta
    // 1) Aggiunge i giocatori mancanti in rosso
    // 2) Modifica i cambi squadra dei giocatori in giallo
    // 3) Rimuove i giocatori andati via
    // Attenzione!! Il tool si aspetta determinate celle in deterinate posizioni. Se tali posizioni cambiano, il tool va aggiornato
    //*
    public class AllignManager : ExcelModifier
    {
        public AllignManager(string fileEliaPath, string filePathFantagazzetta) : base(fileEliaPath, filePathFantagazzetta)
        {
        }

        public override bool Allign(string sheetNameFileElia, string sheetNameFileToCopyFrom)
        {
            Worksheet SheetFantagazzetta = ExcelToCopyFrom.GetSheet(sheetNameFileToCopyFrom);
            Worksheet SheetFileElia = GetSheet(sheetNameFileElia);
            if(AggiungiMancanti(SheetFileElia, SheetFantagazzetta))
            {
                return RimuoviAndatiVia(SheetFileElia, SheetFantagazzetta);
            }
            return false;
        }

        private bool RimuoviAndatiVia(Worksheet SheetFileElia, Worksheet SheetFantagazzetta)
        {
            try
            {
                //Prendo l'ultima riga nn nulla del mio foglio
                Excel.Range last = SheetFileElia.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
                int lastUsedRowMio = last.Row;

                //Prendo l'ultima riga nn nulla del nuovo foglio
                Excel.Range lastSalveMio = SheetFantagazzetta.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
                int lastUsedRow = lastSalveMio.Row;


                for (int i = 2; i <= lastUsedRowMio; i++)
                {
                    bool exsist = false;

                    //Leggo i valori dal mio vecchio file
                    var Name = ((Excel.Range)SheetFileElia.Cells[i, 2]).Value as string ?? "asd";
                    var Ruolo = ((Excel.Range)SheetFileElia.Cells[i, 1]).Value;
                    var Squadra = ((Excel.Range)SheetFileElia.Cells[i, 3]).Value;

                    if (Name == null)
                    {
                        break;
                    }
                    for (int j = 2; j <= lastUsedRow; j++)
                    {
                        //Leggo i vecchi valori dal nuovo file
                        var myName = ((Excel.Range)SheetFantagazzetta.Cells[j, 4]).Value as string ?? "dsa";
                        var myRuolo = ((Excel.Range)SheetFantagazzetta.Cells[j, 2]).Value;
                        var mySquadra = ((Excel.Range)SheetFantagazzetta.Cells[j, 5]).Value;

                        //Confronto
                        if (Name.ToUpper().Equals(myName.ToUpper()))
                        {
                            exsist = true;
                            break;
                        }
                    }

                    //Se vado qui il giocatore non esiste nel nuovo foglio, lo rimuovo
                    if (!exsist)
                    {
                        //Devo Rimuovere dal mio foglio il giocatore (probabilmente è andato via perche esiste nel mio ma nn in quello aggiornato
                        ((Excel.Range)SheetFileElia.Rows[i]).Delete();
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
                Quit();
                ExcelToCopyFrom.CloseBook();
                ExcelToCopyFrom.Quit();
            }
        }
        private bool AggiungiMancanti(Worksheet SheetFileElia, Worksheet SheetFantagazzetta)
        {
            try
            {
                //Prendo l'ultima riga nn nulla del foglio nuovo
                Excel.Range last = SheetFantagazzetta.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
                int lastUsedRow = last.Row;

                //Prendo l'ultima riga nn nulla del mio foglio
                Excel.Range lastSalveMio = SheetFileElia.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
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
                    var Name = ((Excel.Range)SheetFantagazzetta.Cells[i, 4]).Value as string ?? "asd";
                    var Ruolo = ((Excel.Range)SheetFantagazzetta.Cells[i, 2]).Value;
                    var Squadra = ((Excel.Range)SheetFantagazzetta.Cells[i, 5]).Value;
                    var QuotazioneAttuale = ((Excel.Range)SheetFantagazzetta.Cells[i, 6]).Value;

                    for (int j = 2; j <= lastUsedRowMio; j++)
                    {
                        //Leggo i vecchi valori dal mio file
                        string myName = ((Excel.Range)SheetFileElia.Cells[j, 2]).Value as string ?? "dsa";
                        var myRuolo = ((Excel.Range)SheetFileElia.Cells[j, 1]).Value;
                        var mySquadra = ((Excel.Range)SheetFileElia.Cells[j, 3]).Value;

                        //Confronto
                        if (Name.ToUpper().Equals(myName.ToUpper()))
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
                        ((Excel.Range)SheetFileElia.Cells[RowToInsert, 1]).Value = Ruolo;
                        ((Excel.Range)SheetFileElia.Cells[RowToInsert, 2]).Value = Name;
                        ((Excel.Range)SheetFileElia.Cells[RowToInsert, 3]).Value = Squadra;
                        ((Excel.Range)SheetFileElia.Cells[RowToInsert, 4]).Value = QuotazioneAttuale;
                        ((Excel.Range)SheetFileElia.Rows[RowToInsert]).Interior.Color = XlRgbColor.rgbRed;
                        RowToInsert++;
                    }
                    else
                    {
                        if (aggiornaSquadra)
                        {
                            ((Excel.Range)SheetFileElia.Cells[rigaDaAggiornare, 3]).Value = Squadra;
                            ((Excel.Range)SheetFileElia.Cells[rigaDaAggiornare, 4]).Value = QuotazioneAttuale;
                            ((Excel.Range)SheetFileElia.Rows[rigaDaAggiornare]).Interior.Color = XlRgbColor.rgbYellowGreen;
                            rigaDaAggiornare = 0;
                        }
                    }
                }

                //Salvo
                Save();
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                CloseBook();
                ExcelToCopyFrom.CloseBook();
                return false;
            }
        }
    }
}
