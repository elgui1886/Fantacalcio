using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;


namespace ExcelManager
{
    public class ExcelModifier : ExcelReader<ExcelCell>
    {
        public ExcelReader<ReadableCell> ExcelToCopyFrom { get; set; }
        public ExcelModifier(string fileElia, string filePathToCopyFrom) : base(fileElia)
        {
            ExcelToCopyFrom = new ExcelReader<ReadableCell>(filePathToCopyFrom);
        }
        public bool Allign(string sheetNameFileElia, string sheetNameFileToCopyFrom, List<MappingCell> mappingCells)
        {
            var SheetElia = GetSheet(sheetNameFileElia);
            var SheetFantaculo = ExcelToCopyFrom.GetSheet(sheetNameFileToCopyFrom);
            try
            {
                //Prendo l'ultima riga nn nulla del foglio nuovo
                Excel.Range last = SheetFantaculo.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing);
                int lastUsedRow = last.Row;

                //Prendo l'ultima riga nn nulla del mio foglio
                Excel.Range lastSalveMio = SheetElia.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing);
                int lastUsedRowMio = lastSalveMio.Row;

                //Rappresenta la prima riga nulla, ovvero quella dove evenutualmente inseriro i nuovi dati
                int RowToInsert = lastUsedRowMio + 2;

                var _readableCells = mappingCells.Select(c => c.ReadableCell);
                var readableCells = ExcelToCopyFrom.SetColumsIndexesByNames(_readableCells).ToArray();

                var _writableCells = mappingCells.Select(c => c.WritableCell).ToArray();
                var writableCells = SetColumsIndexesByNames(_writableCells).ToArray();

                // Devono essere in egual numero
                if (writableCells.Length != readableCells.Length)
                {
                    throw new Exception("Errore nella configurazione");
                }


                for (int i = 2; i <= lastUsedRow; i++)
                {
                    bool exsist = false;

                    //Rappresenta la riga che dovra essere non inserita, ma aggiornata con la nuova squadra in caso di cambio di maglia
                    int rigaDaAggiornare = 0;

                    foreach (var readableCell in readableCells)
                    {
                        readableCell.Value = ((Excel.Range)SheetFantaculo.Cells[i, readableCell.Index]).Value;
                    }

                    // Essendo il nome il fattore di matching tra il mio file e il file da cui attingere, questo dovrà essere sempre la PRIMA colonna specificata nel file di configurazione
                    var Name = readableCells.First().Value ?? "asd";

                    ////Leggo i valori dal file nuovo

                    for (int j = 2; j <= lastUsedRowMio; j++)
                    {
                        //Leggo i vecchi valori dal mio file
                        var myNameCell = writableCells.First();
                        string myName = ((Excel.Range)SheetElia.Cells[j, myNameCell.Index]).Value as string ?? "asd";


                        //Confronto
                        if (!string.IsNullOrEmpty(Name) && !string.IsNullOrEmpty(myName) && Name.ToUpper().Contains(myName.ToUpper()))
                        {
                            exsist = true;
                            rigaDaAggiornare = j;

                            break;
                        }
                    }

                    if (exsist)
                    {
                        // Skippo il primo poichè è il nome che non voglio sia aggiornato
                        for (int w = 1; w < writableCells.Length; w++)
                        {
                            var writableCell = writableCells[w];
                            var readableCell = readableCells[w];
                            var value = CustomizeValueTool(readableCell);
                            ((Excel.Range)SheetElia.Cells[rigaDaAggiornare, writableCell.Index]).Value = value;
                        }
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

        private object CustomizeValueTool(ReadableCell cell)
        {
            var value = cell.Value.ToString();
            if (cell.ValueFormatter != null)
            {
                value = cell.ValueFormatter(value);
            }

            value = ConvertFromString(value, cell.Type);
            return value;
        }

        private object ConvertFromString(string valoreStringa, string tipoDesiderato)
        {
            tipoDesiderato ??= "string";

            // Ottieni il tipo corrispondente alla stringa specificata
            if (!_typeAlias.TryGetValue(tipoDesiderato, out Type tipo))
            {
                return valoreStringa;
            }
            // Converte la stringa nel tipo desiderato utilizzando il metodo Parse
            return Convert.ChangeType(valoreStringa, tipo);

        }
    }
}
