using ExcelManager;
using System;

namespace FantaGoat
{
    internal class Program
    {
        public const string Portieri = "Portieri";
        public const string Difensori = "Difensori";
        public const string Centrocampisti = "Centrocampisti";
        public const string Attaccanti = "Attaccanti";

        public const string MioFilePath = @"C:\Users\eliag\Desktop\Elia\FantaLista\2023-2024\EG_ListoneAsta_2023-2024.xlsx";
        public const string FileFantaLabPath = "C:\\Users\\eliag\\Desktop\\Elia\\FantaLista\\Fantacalcio\\AllineaPrezziFantaLab\\src\\Strategia Il Profeta.xlsx";
        public static string[] columnNameToRead = { "Nome", "Fascia", "Prezzo" };
        public static string[] columnNameToWrite = { "Nome", "SLOT PROFETA", "PREZZO PROFETA" };



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

            ExcelModifier manager = new(MioFilePath, FileFantaLabPath);
            if (manager.Allign(sheetname, columnNameToWrite, arg, columnNameToRead, Tool.Fantalab))
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