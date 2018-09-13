using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;


namespace ExeCasoTeste
{
    class Program
    {
        
        private static string logLevel = "ERRO";

        static void Main(string[] args)
        {
            try
            {
                logLevel = args[0];
            }
            catch (Exception)
            {
            }

            string path = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);

            DirectoryInfo d = new DirectoryInfo(path + "\\casoteste\\");
            FileInfo[] Files = d.GetFiles("*.xlsx");
            Parallel.ForEach(Files, (file) =>
                {
                    CasoTeste caso = new CasoTeste();
                    caso.arquivo = file.FullName;
                    caso.logLevel = logLevel;
                    caso.ProcessarCasoTeste();
                });
        }

        
    }
}