using System;
using System.IO;

namespace ExeCasoTeste
{
    public class Log
    {
        int countErro = 0;
        string logLevel = "INFO";
        string arquivo = "";

        public Log(string _logLevel, string _arquivo)
        {
            logLevel = _logLevel;
            arquivo = _arquivo;
        }

        public void Registrar(string texto, string tipo = "INFO")
        {
            string data = "[" + DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToShortTimeString() + ":" + DateTime.Now.Second.ToString().PadLeft(2, '0') + "] ";

            if (tipo == "INFO")
            {
                if (logLevel != "ERRO")
                {
                    Console.Write(data + Path.GetFileNameWithoutExtension(arquivo) + "-" + texto);
                    File.AppendAllText(arquivo, data + texto);
                }
            }
            else
            {
                Console.Write(data + Path.GetFileNameWithoutExtension(arquivo) + "-" + texto);
                File.AppendAllText(arquivo, data + texto);
                countErro++;
            }
        }

        public int getCountErro()
        {
            return countErro;
        }
    }
}
