using ExcelManager;
using System;
using System.Collections.Generic;

namespace AllineaPrezziFantaculo
{
    internal class Program
    {
        public const string Portieri = "Portieri";
        public const string Difensori = "Difensori";
        public const string Centrocampisti = "Centrocampisti";
        public const string Attaccanti = "Attaccanti";

        public const string MioFilePath = @"C:\Dev\FantaLista\2024-2025\EG_ListoneAsta_2024-2025.xlsx";
        public const string FileFantaculoPath = @"C:\Users\eliag\Downloads\Listone_Fantaculo.xlsx";


        public static List<MappingCell> mapping = new()
        {
            new MappingCell { WritableCell = new ExcelCell { Name = "Nome" }, ReadableCell = new ReadableCell { Name = "name" } },
            new MappingCell { WritableCell = new ExcelCell { Name = "FASCIA FC" }, ReadableCell = new ReadableCell { Name = "slot", Type = "double" } },
            new MappingCell { WritableCell = new ExcelCell { Name = "PREZZO FC" }, ReadableCell = new ReadableCell { Name = "pfc", Type = "double" } },
            new MappingCell { WritableCell = new ExcelCell { Name = "FC_PMA" }, ReadableCell = new ReadableCell { Name = "pma", Type = "double" }}
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

            ExcelModifier manager = new(MioFilePath, FileFantaculoPath);
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