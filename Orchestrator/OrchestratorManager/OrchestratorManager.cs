using AllineaListone;
using AllineaPrezziFantaculo;
using AllineaPrezziFantaLab;
using CostoAstaPrecedente;
using FantaGoat;
using System;


namespace Orchestrator.OrchestratorManager
{
    public static class OrchestratorManager
    {
        public static bool AllineaListone(string mioFilePath, string mioSheet, string sheetFileToCopy)
        {
            #region AllineaListone
            string listonePath = "C:\\Users\\Elia\\Desktop\\Elia\\FantaLista\\2023-2024\\Quotazioni_Fantacalcio_Stagione_2023_24.xlsx";
            AllignManager AllignManager = new(mioFilePath, listonePath);
            if (AllignManager.Allign(mioSheet, sheetFileToCopy))
            {
                Console.WriteLine("Fatto allinea listone, tutto ok!");
                return true;
            }
            else
            {
                Console.WriteLine("Qualcosa è andato storto in  allinea listone");
                return false;
            }
            #endregion
        }

        public static bool AllineaFantaculo(string mioFilePath, string mioSheet, string sheetFileToCopy) 
        {
            #region AllineaFantaculo
            string fileFantaculoPath = "C:\\Users\\Elia\\Downloads\\Listone_Fantaculo.xlsx";
            FantaculoManager FantaculoManager = new(mioFilePath, fileFantaculoPath);
            if (FantaculoManager.Allign(mioSheet, sheetFileToCopy))
            {
                Console.WriteLine("Fatto fantaculo, tutto ok!");
                return true;
            }
            else
            {
                Console.WriteLine("Qualcosa è andato storto in fantaculo");
                return false;
            }
            #endregion
        }

        public static bool AllineaFantalab(string mioFilePath, string mioSheet, string sheetFileToCopy)
        {
            #region AllineaFantaLab
            string fileFantalabPath = @"C:\Users\Elia\Downloads\StrategiaProfeta.xlsx";
            FantaLabManager FantaLabManager = new(mioFilePath, fileFantalabPath);
            if (FantaLabManager.Allign(mioSheet, sheetFileToCopy))
            {
                Console.WriteLine("Fatto fantalab, tutto ok!");
                return true;
            }
            else
            {
                Console.WriteLine("Qualcosa è andato storto in fantalab");
                return false;
            }
            #endregion
        }

        public static bool AllineaFantaGoat(string mioFilePath, string mioSheet)
        {
            #region AllineaFantaGoat
            string fileFantaGoatPath = @"C:\Users\Elia\Downloads\lega_Slot_ Lega a 10 partecipanti.xlsx";
            FantaGoatManager FantaGoatManager = new(mioFilePath, fileFantaGoatPath);
            if (FantaGoatManager.Allign(mioSheet, "Sheet1"))
            {
                Console.WriteLine("Fatto fantagoat, tutto ok!");
                return true;
            }
            else
            {
                return false;
            }
            #endregion
        }

        public static bool AllineaCostoAstaPrecedente(string mioFilePath, string mioSheet)
        {
            #region AllineaCostoAstaPrecedente
            string fileCostoAstaPrecedentePath = @"C:\Users\Elia\Desktop\Elia\FantaLista\2022-2023\a.xlsx";

            CostoAstaPrecedenteManager CostoAstaPrecedenteManager = new(mioFilePath, fileCostoAstaPrecedentePath);


            if (CostoAstaPrecedenteManager.Allign(mioSheet, "a"))
            {
                Console.WriteLine("Fatto, tutto ok!");
                return true;
            }
            else
            {
                Console.WriteLine("Qualcosa è andato storto ");
                return false;
            }
            #endregion
        }
    }



}
