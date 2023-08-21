using ExcelManager;
using System;

namespace AllineaPrezziFantaculo
{
    internal class Program
    {
        public const string Portieri = "Portieri";
        public const string Difensori = "Difensori";
        public const string Centrocampisti = "Centrocampisti";
        public const string Attaccanti = "Attaccanti";

        public const string MioFilePath = @"C:\Users\eliag\Desktop\Elia\FantaLista\2023-2024\EG_ListoneAsta_2023-2024.xlsx";
        public const string FileFantaculoPath = @"C:\Dev\Fantacalcio\AllineaPrezziFantaculo\src\Listone_Fantaculo.xlsx";
        public static string[] columnNameToRead = { "name", "slot", "pfc", "pma" };
        public static string[] columnNameToWrite = { "Nome", "SLOT FC", "PREZZO FC", "PREZZO ASTA" };



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

            ExcelModifier manager = new(MioFilePath, FileFantaculoPath);
            if (manager.Allign(sheetname, columnNameToWrite, arg, columnNameToRead, Tool.Fantaculo))
            {
                Console.WriteLine("Fatto, tutto ok!");
            }
            else
            {
                Console.WriteLine("Qualcosa è andato storto ");
                Console.ReadLine();
            }
        }
    }
}