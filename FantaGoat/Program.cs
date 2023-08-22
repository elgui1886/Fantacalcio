using ExcelManager;
using System;
using System.Collections.Generic;

namespace FantaGoat
{
    internal class Program
    {
        public const string Portieri = "Portieri";
        public const string Difensori = "Difensori";
        public const string Centrocampisti = "Centrocampisti";
        public const string Attaccanti = "Attaccanti";

        public const string MioFilePath = @"C:\Users\eliag\Desktop\Elia\FantaLista\2023-2024\EG_ListoneAsta_2023-2024.xlsx";
        public const string FileFantaGoatPath = "C:\\Dev\\Fantacalcio\\FantaGoat\\src\\lega_Slot_ Lega a 10 partecipanti.xlsx";


        public static List<MappingCell> mapping = new()
        {
            new MappingCell { WritableCell = new ExcelCell { Name = "Nome" }, ReadableCell = new ReadableCell { Name = "Player" } },
            new MappingCell { WritableCell = new ExcelCell { Name = "SLOT FG" }, ReadableCell = new ReadableCell { Name = "Slot", Type = "double", ValueFormatter = slot => slot.Replace("° SLOT", "") } },
            new MappingCell { WritableCell = new ExcelCell { Name = "PREZZO FANTAGOAT" }, ReadableCell = new ReadableCell { Name = "Prezzo massimo", Type = "double" } },
        };



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

            ExcelModifier manager = new(MioFilePath, FileFantaGoatPath);
            if (manager.Allign(sheetname, "Sheet1", mapping))
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