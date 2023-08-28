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

        public const string MioFilePath = @"C:\Dev\FantaLista\2023-2024\EG_ListoneAsta_2023-2024.xlsx";
        public const string FileFantaLabPath = "C:\\Users\\eliag\\Downloads\\Strategia Il Profeta.xlsx";

        public static List<MappingCell> mapping = new()
        {
            new MappingCell { WritableCell = new ExcelCell { Name = "Nome" }, ReadableCell = new ReadableCell { Name = "Nome" } },
            new MappingCell { WritableCell = new ExcelCell { Name = "SLOT PROFETA" }, ReadableCell = new ReadableCell { Name = "Fascia", ValueFormatter = (slot) =>
            {
                    if (slot == "Top")
                    {
                        slot = "1";
                    }
                    else if (slot == "Semi-Top")
                    {
                        slot = "2";
                    }
                    else if (slot == "Terza Fascia")
                    {
                        slot = "3";
                    }
                    return slot;

            } } },
            new MappingCell { WritableCell = new ExcelCell { Name = "PREZZO PROFETA" }, ReadableCell = new ReadableCell { Name = "Prezzo", Type = "double" } },
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

            ExcelModifier manager = new(MioFilePath, FileFantaLabPath);
            if (manager.Allign(sheetname, arg, mapping))
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