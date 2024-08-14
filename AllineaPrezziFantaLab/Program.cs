using ExcelManager;
using System;
using System.Collections.Generic;

namespace FantaLab
{
    /*
    * Il profeta
    */
    internal class Program
    {
        public const string Portieri = "Portieri";
        public const string Difensori = "Difensori";
        public const string Centrocampisti = "Centrocampisti";
        public const string Attaccanti = "Attaccanti";

        public const string MioFilePath = @"C:\Dev\FantaLista\2024-2025\EG_ListoneAsta_2024-2025.xlsx";



        public const string FileFantaLabPath = "C:\\Users\\eliag\\Downloads\\Strategia CarmySpecial.xlsx";
        public const string Nomeprezzo = "PREZZO CARMY";
        public const string Nomefascia = "FASCIA CARMY";
        //public const string FileFantaLabPath = "C:\\Users\\eliag\\Downloads\\Strategia Lorenzo Cantarini.xlsx";
        //public const string Nomeprezzo = "PREZZO CANTARINI";
        //public const string Nomefascia = "FASCIA CANTARINI";
        //public const string FileFantaLabPath = "C:\\Users\\eliag\\Downloads\\Strategia Recosta.xlsx";
        //public const string Nomeprezzo = "PREZZO RECOSTA";
        //public const string Nomefascia = "FASCIA RECOSTA";
        //public const string FileFantaLabPath = "C:\\Users\\eliag\\Downloads\\Strategia Walk Off Sport.xlsx";
        //public const string Nomeprezzo = "PREZZO WALK";
        //public const string Nomefascia = "FASCIA WALK";
        //public const string FileFantaLabPath = "C:\\Users\\eliag\\Downloads\\Strategia Fanta__Boom.xlsx";
        //public const string Nomeprezzo = "PREZZO FANTABOOM";
        //public const string Nomefascia = "FASCIA FANTABOOM";
        //public const string FileFantaLabPath = "C:\\Users\\eliag\\Downloads\\Strategia Il Profeta.xlsx";
        //public const string Nomeprezzo = "PREZZO PROFETA";
        //public const string Nomefascia = "FASCIA PROFETA";
        //public const string FileFantaLabPath = "C:\\Users\\eliag\\Downloads\\Strategia SOS Fanta";
        //public const string Nomeprezzo = "PREZZO SOS";
        //public const string Nomefascia = "FASCIA SOS";
        //public const string FileFantaLabPath = "C:\\Users\\eliag\\Downloads\\Strategia Luca Diddi _ Il Tattico";
        //public const string Nomeprezzo = "PREZZO TATTICO";
        //public const string Nomefascia = "FASCIA TATTICO";


        public static List<MappingCell> mapping = new()
        {
            new MappingCell { WritableCell = new ExcelCell { Name = "Nome" }, ReadableCell = new ReadableCell { Name = "Nome" } },
            new MappingCell { WritableCell = new ExcelCell { Name = Nomeprezzo }, ReadableCell = new ReadableCell { Name = "Prezzo", Type = "double" } },
            new MappingCell { WritableCell = new ExcelCell { Name = "PMAFL" }, ReadableCell = new ReadableCell { Name = "PMA" } },
            new MappingCell { WritableCell = new ExcelCell { Name = Nomefascia }, ReadableCell = new ReadableCell { Name = "Fascia", ValueFormatter = slot => {
                if(slot is null)
                {
                    return string.Empty;
                }
                if(slot.Contains("Non Impostata"))
                {
                    return string.Empty;
                }
                if(slot.Contains("°") )
                {
                    return slot.Replace("°", " ");
                }
                return slot;
            } } },
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