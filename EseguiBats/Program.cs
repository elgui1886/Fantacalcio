using System.Diagnostics;

namespace EseguiBats
{
    internal class Program
    {
        static void Main(string[] args)
        {
            // Specifica il percorso della cartella
            string folderPath = @"C:\Dev\Fantacalcio\build";

            // Controlla se la cartella esiste
            if (!Directory.Exists(folderPath))
            {
                Console.WriteLine("La cartella specificata non esiste.");
                Console.ReadLine();
                return;
            }

            // Ottiene tutti i file .bat nella cartella
            string[] batFiles = Directory.GetFiles(folderPath, "*.bat");

            // Se non ci sono file .bat, esci dal programma
            if (batFiles.Length == 0)
            {
                Console.WriteLine("Non ci sono file .bat nella cartella specificata.");
                Console.ReadLine();
                return;
            }

            // Esegui ogni file .bat uno alla volta
            foreach (string batFile in batFiles)
            {
                Console.WriteLine($"Esecuzione del file: {Path.GetFileName(batFile)}");

                try
                {
                    // Configura il processo per eseguire il file .bat
                    ProcessStartInfo processInfo = new ProcessStartInfo();
                    processInfo.FileName = batFile;
                    processInfo.FileName = "cmd.exe";
                    processInfo.Arguments = $"/c \"{batFile}\"";
                    processInfo.RedirectStandardOutput = true;
                    processInfo.RedirectStandardError = true;
                    processInfo.UseShellExecute = false;
                    processInfo.CreateNoWindow = false;

                    using (Process process = Process.Start(processInfo))
                    {
                        // Leggi l'output del processo
                        string output = process.StandardOutput.ReadToEnd();
                        string error = process.StandardError.ReadToEnd();

                        process.Close();
                        process.Dispose();

                        // Stampa l'output del processo
                        Console.WriteLine(output);

                        // Se ci sono errori, stampali
                        if (!string.IsNullOrEmpty(error))
                        {
                            Console.WriteLine("Errori:");
                            Console.WriteLine(error);
                            Console.ReadLine();
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Errore durante l'esecuzione del file {Path.GetFileName(batFile)}: {ex.Message}");
                    Console.ReadLine();
                }

                Console.WriteLine(); // Righe vuote per separare i risultati tra i file
            }

            Console.WriteLine("Esecuzione completata.");
            Console.ReadLine();
        }
    }
}
