using ConfrontoListaFanta.ConfrontaFile;
using System;

namespace ConfrontoListaFanta
{
    class Program
    {
        public const string Portieri = "Portieri";
        public const string Difensori = "Difensori";
        public const string Centrocampisti = "Centrocampisti";
        public const string Attaccanti = "Attaccanti";

        static void Main(string[] args)
        {
            string sheetname = string.Empty;
            var arg = args[0].ToUpper();
            switch (arg)
            {
                case "P":
                    sheetname = Portieri;
                    break;
                case "D":
                    sheetname = Difensori;
                    break;
                case "C":
                    sheetname = Centrocampisti;
                    break;
                case "A":
                    sheetname = Attaccanti;
                    break;
                default:
                    break;
            }

            Console.WriteLine("Allineo: " + sheetname);
            if (string.IsNullOrEmpty(sheetname))
            {
                Console.WriteLine("Nessun argomento passato");
                return;
            }
            string mioFilePath = "C:\\Users\\Elia\\Desktop\\Elia\\FantaLista\\2023-2024\\EG_ListoneAsta_2023-2024.xlsx";
            string nuovoFilePath = "C:\\Users\\Elia\\Desktop\\Elia\\FantaLista\\2023-2024\\Quotazioni_Fantacalcio_Stagione_2023_24.xlsx";
            string fileFantaculoPath = "C:\\Users\\Elia\\Downloads\\Listone_Fantaculo.xlsx";


            //Confronter c = new Confronter();
            //Microsoft.Office.Interop.Excel.Worksheet mioFileCommentato = c.ReadFileExcelMio(mioFilePath, sheetname);
            //Microsoft.Office.Interop.Excel.Worksheet nuovoFileAggiornato = c.ReadListone(nuovoFilePath, sheetname);
            //if (!c.AggiungiMancanti(nuovoFileAggiornato, mioFileCommentato))
            //{
            //    Console.WriteLine("Qualcosa è andato storto nella aggiunta dei mancanti");
            //    Console.ReadLine();
            //}
            //else
            //{
            //    if (!c.RimuoviAndatiVia(mioFileCommentato, nuovoFileAggiornato))
            //    {
            //        Console.WriteLine("Qualcosa è andato storto nella rimmozione degli andati via");
            //        Console.ReadLine();
            //    }
            //    else
            //    {
            //        Console.WriteLine("Fatto, tutto ok!");
            //        Console.ReadLine();
            //    }
            //}

            AllineaPrezzi a = new AllineaPrezzi();
            Microsoft.Office.Interop.Excel.Worksheet mioFileCommentato2 = a.ReadFileExcelMio(mioFilePath, sheetname);
            Microsoft.Office.Interop.Excel.Worksheet fileFantaculo = a.ReadFantaculo(fileFantaculoPath, arg);
            if (!a.AllineaPrezziESlot(fileFantaculo, mioFileCommentato2))
            {
                Console.WriteLine("Qualcosa è andato storto nell aggiornamento di prezzi e slot");
                Console.ReadLine();
            }
            else
            {
                Console.WriteLine("Fatto, tutto ok!");
                Console.ReadLine();
            }
        }
    }
}
