using AllineaListone;
using AllineaPrezziFantaculo;
using AllineaPrezziFantaLab;
using CostoAstaPrecedente;
using FantaGoat;
using System;

namespace Orchestrator
{
    internal class Program
    {
        public const string Portieri = "Portieri";
        public const string Difensori = "Difensori";
        public const string Centrocampisti = "Centrocampisti";
        public const string Attaccanti = "Attaccanti";

        static void Main(string[] args)
        {
            string mioFilePath = "C:\\Users\\Elia\\Desktop\\Elia\\FantaLista\\2023-2024\\EG_ListoneAsta_2023-2024.xlsx";


            #region AllineaListone
            //Console.WriteLine("AllineaListone");
            //OrchestratorManager.OrchestratorManager.AllineaListone(mioFilePath, Portieri, Portieri);
            //OrchestratorManager.OrchestratorManager.AllineaListone(mioFilePath, Difensori, Difensori);
            //OrchestratorManager.OrchestratorManager.AllineaListone(mioFilePath, Centrocampisti, Centrocampisti);
            //OrchestratorManager.OrchestratorManager.AllineaListone(mioFilePath, Attaccanti, Attaccanti);
            #endregion

            #region AllineaFantaculo
            Console.WriteLine("AllineaFantaculo");
            OrchestratorManager.OrchestratorManager.AllineaFantaculo(mioFilePath, Portieri, "P");
            OrchestratorManager.OrchestratorManager.AllineaFantaculo(mioFilePath, Difensori, "D");
            OrchestratorManager.OrchestratorManager.AllineaFantaculo(mioFilePath, Centrocampisti, "C");
            OrchestratorManager.OrchestratorManager.AllineaFantaculo(mioFilePath, Attaccanti, "A");
            #endregion

            #region AllineaFantaLab
            Console.WriteLine("AllineaFantaLab");
            OrchestratorManager.OrchestratorManager.AllineaFantalab(mioFilePath, Portieri, "P");
            OrchestratorManager.OrchestratorManager.AllineaFantalab(mioFilePath, Difensori, "D");
            OrchestratorManager.OrchestratorManager.AllineaFantalab(mioFilePath, Centrocampisti, "C");
            OrchestratorManager.OrchestratorManager.AllineaFantalab(mioFilePath, Attaccanti, "A");
            #endregion

            #region AllineaFantaGoat
            Console.WriteLine("AllineaFantaGoat");
            OrchestratorManager.OrchestratorManager.AllineaFantaGoat(mioFilePath, Portieri);
            OrchestratorManager.OrchestratorManager.AllineaFantaGoat(mioFilePath, Difensori);
            OrchestratorManager.OrchestratorManager.AllineaFantaGoat(mioFilePath, Centrocampisti);
            OrchestratorManager.OrchestratorManager.AllineaFantaGoat(mioFilePath, Attaccanti);
            #endregion

            #region AllineaCostoAstaPrecedente
            Console.WriteLine("AllineaCostoAstaPrecedente");
            OrchestratorManager.OrchestratorManager.AllineaCostoAstaPrecedente(mioFilePath, Portieri);
            OrchestratorManager.OrchestratorManager.AllineaCostoAstaPrecedente(mioFilePath, Difensori);
            OrchestratorManager.OrchestratorManager.AllineaCostoAstaPrecedente(mioFilePath, Centrocampisti);
            OrchestratorManager.OrchestratorManager.AllineaCostoAstaPrecedente(mioFilePath, Attaccanti);
            #endregion
        }
    }

}
