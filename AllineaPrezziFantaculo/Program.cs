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
            string fileFantaculoPath = "C:\\Users\\Elia\\Downloads\\Listone_Fantaculo.xlsx";


            ExcelReader fileToCopyFrom = new ExcelReader(fileFantaculoPath);
            FantaculoManager excelModifier = new FantaculoManager(mioFilePath, fileToCopyFrom);


            if (excelModifier.Allign(sheetname, arg))
            {
                Console.WriteLine("Fatto, tutto ok!");
                Console.ReadLine();
            }
            else
            {
                Console.WriteLine("Qualcosa è andato storto ");
                Console.ReadLine();
            }
        }
    }
}