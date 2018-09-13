using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExeCasoTeste
{
    public class CasoTeste
    {
        public string arquivo { get; set; }

        public string logLevel { get; set; }

        private List<string> listaStrings = new List<string>();

        private Oracle bd = new Oracle("");

        public void ProcessarCasoTeste()
        {
            Dictionary<string, string> valoresDeletar = new Dictionary<string, string>();
            Dictionary<string, string> valoresInserir = new Dictionary<string, string>();
            Dictionary<string, string> valoresExecutar = new Dictionary<string, string>();
            Dictionary<string, string> valoresSaida = new Dictionary<string, string>();
            Dictionary<string, string> valoresVerificar = new Dictionary<string, string>();
            Dictionary<string, string> valoresScript = new Dictionary<string, string>();

            string arquivoLog = Path.GetDirectoryName(arquivo).Replace("\\casoteste", "") + "\\log\\" + Path.ChangeExtension(Path.GetFileName(arquivo), ".log");
            string pathScript = Path.GetDirectoryName(arquivo).Replace("\\casoteste", "") + "\\script";

            Log log = new Log(logLevel, arquivoLog);

            if (!File.Exists(arquivo))
            {
                log.Registrar("Arquivo " + arquivo + " não existe.", "ERRO");
                return;
            }

            if (!File.Exists(arquivoLog))
            {
                FileStream f = File.Create(arquivoLog);
                f.Close();
            }
            else
            {
                File.Delete(arquivoLog);
                FileStream f = File.Create(arquivoLog);
                f.Close();
            }

            log.Registrar("### Validando arquivo de strings" + Environment.NewLine);
            string arquivoStrings = Path.GetDirectoryName(arquivoLog) + "\\string.txt";
            if (File.Exists(arquivoStrings))
            {
                listaStrings = File.ReadAllText(arquivoStrings).Replace("\n", string.Empty).Trim().Split('\r').ToList();
                log.Registrar("### Arquivo de strings encontrado: " + listaStrings.Count.ToString() + " registro(s)" + Environment.NewLine);
            }
            else
            {
                log.Registrar("### Arquivo de strings não encontrado!" + Environment.NewLine);
            }

            log.Registrar("### Validando conexão com o banco" + Environment.NewLine);

            if (!File.Exists("conexao.txt"))
            {
                log.Registrar("### Arquivo de conexão não encontrado!" + Environment.NewLine);
                return;
            }

            string conex = File.ReadAllText("conexao.txt");

            bd = new Oracle(conex);
            string teste = bd.Testa();

            if (teste == "OK")
            {
                log.Registrar("### Conexão com o banco de dados ok!" + Environment.NewLine);
            }
            else
            {
                log.Registrar("### Erro ao conectar ao banco de dados: " + teste + Environment.NewLine, "ERRO");
                return;
            }

            log.Registrar("" + Environment.NewLine);

            log.Registrar("### Iniciando processamento do arquivo excel" + Environment.NewLine);
            processarArquivoExcel(log, arquivo,
                ref valoresDeletar,
                ref valoresInserir,
                ref valoresExecutar,
                ref valoresSaida,
                ref valoresVerificar,
                ref valoresScript);

            log.Registrar("### Fim do processamento do arquivo excel" + Environment.NewLine);

            log.Registrar("" + Environment.NewLine);

            if (valoresDeletar.Count > 0)
            {
                log.Registrar("### Iniciando deleção de valores: " + valoresDeletar.Count.ToString() + " registro(s)" + Environment.NewLine);
                deletarValores(log, valoresDeletar);
                log.Registrar("### Finalizando deleção de valores" + Environment.NewLine);
                log.Registrar("" + Environment.NewLine);
            }

            if (valoresScript.Count > 0)
            {
                log.Registrar("### Iniciando execução de scripts: " + valoresScript.Count.ToString() + " registro(s)" + Environment.NewLine);
                ExecutarScripts(log, pathScript, valoresScript);
                log.Registrar("### Finalizando execução de scripts" + Environment.NewLine);
                log.Registrar("" + Environment.NewLine);
            }

            if (valoresInserir.Count > 0)
            {
                log.Registrar("### Iniciando insert/update de valores: " + valoresInserir.Count.ToString() + " registro(s)" + Environment.NewLine);
                upsertValores(log, valoresInserir);
                log.Registrar("### Finalizando insert/update de valores" + Environment.NewLine);
                log.Registrar("" + Environment.NewLine);
            }

            if (valoresExecutar.Count > 0)
            {
                executarTestes(log, valoresExecutar, valoresSaida);
                log.Registrar("### Finalizando execução de objetos" + Environment.NewLine);
                log.Registrar("" + Environment.NewLine);
            }

            if (valoresVerificar.Count > 0)
            {
                log.Registrar("### Iniciando verificação de valores: " + valoresVerificar.Count.ToString() + " registro(s)" + Environment.NewLine);
                VerificarValores(log, valoresVerificar);
                log.Registrar("### Finalizando verificação de valores" + Environment.NewLine);
                log.Registrar("" + Environment.NewLine);
            }

            log.Registrar(">>> Verificações finalizadas. " + log.getCountErro().ToString() + " erro(s) encontrados." + Environment.NewLine, "ERRO");
        }

        private bool CampoSempreString(string campo)
        {
            return listaStrings.Any(campo.ToUpper().Contains);
        }

        private void VerificarValores(Log log, Dictionary<string, string> valoresVerificar)
        {
            List<string> tabelas = new List<string>();

            foreach (var item in valoresVerificar)
            {
                string nomeTabela = item.Key.Split('.')[0].Replace("*", "");
                if (!tabelas.Contains(nomeTabela))
                {
                    tabelas.Add(nomeTabela);
                }
            }

            foreach (var item in tabelas)
            {
                string where = "";
                string campos = "";

                log.Registrar("# Processando tabela  " + item + Environment.NewLine);
                string comando = "select {2} from {0} where {1}";

                Dictionary<string, string> verificar = new Dictionary<string, string>();

                foreach (var item2 in valoresVerificar)
                {
                    string tabela = item2.Key.Split('.')[0].Replace("*", "");
                    string campo = item2.Key.Split('.')[1];

                    if (tabela == item)
                    {
                        if (item2.Key.Contains("*" + item + "."))
                        {

                            string valor = item2.Value;

                            if (CampoSempreString(campo))
                            {
                                valor = "'" + valor + "'";
                            }
                            else
                            {
                                if (isDateTime(valor))
                                {
                                    valor = returnDateTime(valor);
                                }
                                else if (isDouble(valor))
                                {
                                    valor = valor.Replace(",", ".");
                                }
                                else
                                {
                                    valor = "'" + valor + "'";
                                }
                            }

                            if (String.IsNullOrEmpty(where))
                            {
                                where += campo + " = " + valor;
                            }
                            else
                            {
                                where += " and " + campo + " = " + valor;
                            }
                        }
                        else
                        {
                            if (String.IsNullOrEmpty(campos))
                            {
                                campos += campo;
                            }
                            else
                            {
                                campos += ", " + campo;
                            }

                            verificar.Add(campo, item2.Value);
                        }
                    }
                }

                comando = String.Format(comando, item, where, campos);


                System.Data.DataTable dt = new System.Data.DataTable();

                try
                {
                    dt = bd.ExecutaComando(comando);
                }
                catch (Exception Ex)
                {
                    log.Registrar("Comando: " + comando + Environment.NewLine, "ERRO");
                    log.Registrar("Erro ao executar comando: " + Ex.Message + Environment.NewLine, "ERRO");
                }

                foreach (System.Data.DataRow linha in dt.Rows)
                {
                    foreach (var itemVerificar in verificar)
                    {
                        string valorBanco = linha[itemVerificar.Key.ToString()].ToString();
                        string valorPlanilha = itemVerificar.Value;

                        if (valorBanco != valorPlanilha)
                        {
                            log.Registrar("Erro ao comparar resultados da tabela " + item + ", campo " + itemVerificar.Key + ": Valor no Banco " + valorBanco + " Valor Planilha " + valorPlanilha + Environment.NewLine, "ERRO");
                        }
                    }
                }
            }

        }

        private void executarTestes(Log log, Dictionary<string, string> valoresExecutar, Dictionary<string, string> valoresSaida)
        {
            Dictionary<int, string> testes = new Dictionary<int, string>();

            foreach (var item in valoresExecutar)
            {
                string[] reg = item.Key.Split('|');
                if (!testes.ContainsKey(Convert.ToInt32(reg[0])))
                {
                    testes.Add(Convert.ToInt32(reg[0]), reg[1]);
                }
            }

            log.Registrar("### Iniciando execução de objetos: " + testes.Count.ToString() + " registro(s)" + Environment.NewLine);

            foreach (var item in testes)
            {
                log.Registrar("Preparando parâmetros para execução do teste " + item.Key.ToString() + ": " + item.Value + Environment.NewLine);

                Dictionary<string, string> paramEntrada = new Dictionary<string, string>();

                // parametros de entrada
                foreach (var itemEntrada in valoresExecutar)
                {
                    if (itemEntrada.Key.StartsWith(item.Key.ToString() + "|" + item.Value))
                    {
                        string parametro = itemEntrada.Key.Replace(item.Key.ToString() + "|" + item.Value + "|", "");
                        string valor = itemEntrada.Value;

                        if (valor.ToUpper() == "TRUE" || valor.ToUpper() == "FALSE")
                        {
                            paramEntrada.Add(parametro, "B|" + valor);
                        }
                        else if (CampoSempreString(parametro))
                        {
                            paramEntrada.Add(parametro, "V|" + valor);
                        }
                        else
                        {
                            if (isDateTime(valor))
                            {
                                paramEntrada.Add(parametro, "D|" + returnDateTime(valor));
                            }
                            else if (isDouble(valor))
                            {
                                paramEntrada.Add(parametro, "F|" + valor.Replace(",", "."));
                            }
                            else if (isInteger(valor))
                            {
                                paramEntrada.Add(parametro, "I|" + valor);
                            }
                            else
                            {
                                paramEntrada.Add(parametro, "V|" + valor);
                            }
                        }
                    }
                }

                Dictionary<string, string> paramSaida = new Dictionary<string, string>();

                foreach (var itemSaida in valoresSaida)
                {
                    if (itemSaida.Key.ToString().StartsWith(item.Key.ToString() + "|"))
                    {
                        string parametro = itemSaida.Key.Replace(item.Key.ToString() + "|", "");
                        string valor = itemSaida.Value;

                        if (CampoSempreString(parametro))
                        {
                            paramSaida.Add(parametro, "V|" + valor);
                        }
                        else
                        {
                            if (isDateTime(valor))
                            {
                                if (parametro.ToUpper() == "RESULT")
                                {
                                    paramSaida.Add(parametro, "D|" + returnDateTime(valor));
                                }
                                else
                                {
                                    paramSaida.Add(parametro, "*D|" + returnDateTime(valor));
                                }
                            }
                            else if (isDouble(valor))
                            {
                                if (parametro.ToUpper() == "RESULT")
                                {
                                    paramSaida.Add(parametro, "*F|" + valor.Replace(",", "."));
                                }
                                else
                                {
                                    paramSaida.Add(parametro, "F|" + valor.Replace(",", "."));
                                }

                            }
                            else if (isInteger(valor))
                            {
                                if (parametro.ToUpper() == "RESULT")
                                {
                                    paramSaida.Add(parametro, "*I|" + valor);
                                }
                                else
                                {
                                    paramSaida.Add(parametro, "I|" + valor);
                                }
                            }
                            else
                            {
                                if (parametro.ToUpper() == "RESULT")
                                {
                                    paramSaida.Add(parametro, "*V|" + valor);
                                }
                                else
                                {
                                    paramSaida.Add(parametro, "V|" + valor);
                                }
                            }
                        }
                    }
                }

                Dictionary<string, string> ret = new Dictionary<string, string>();

                try
                {
                    log.Registrar("Iniciando execução da procedure: " + item.Value + Environment.NewLine);
                    ret = bd.TesteObjetoBanco(item.Value, paramEntrada, paramSaida);
                    log.Registrar("Procedure executada com sucesso!" + Environment.NewLine);
                }
                catch (Exception Ex)
                {
                    log.Registrar("Erro ao executar procedure " + item.Value + ": " + Ex.Message + Environment.NewLine, "ERRO");
                }

                log.Registrar("Iniciando validação das saídas no teste " + item.Value + Environment.NewLine);
                foreach (var itemRetorno in ret)
                {
                    string valor = itemRetorno.Value;
                    string valorPlanilha = paramSaida[itemRetorno.Key].Split('|')[1];

                    if (valor != valorPlanilha)
                    {
                        log.Registrar("Erro ao comparar resultados da procedure " + item.Value + ", parâmetro " + itemRetorno.Key + ": Valor Gerado " + valor + " Valor Planilha " + valorPlanilha + Environment.NewLine, "ERRO");
                    }
                }
                log.Registrar("Fim da validação das saídas" + Environment.NewLine);
            }
        }

        private void ExecutarScripts(Log log, string path, Dictionary<string, string> valoresScript)
        {
            var scripts = valoresScript
                                .GroupBy(u => u.Key.Split('#')[0])
                                .Select(j => j.Key.Split('#')[0])
                                .ToList();

            foreach (var arquivoScript in scripts)
            {
                string arquivo = path + "\\" + arquivoScript;

                if (!File.Exists(arquivo))
                {
                    log.Registrar("Arquivo de script " + arquivo + " não encontrado." + Environment.NewLine, "ERRO");
                }
                else
                {
                    string script = File.ReadAllText(arquivo);

                    foreach (var item in valoresScript.Where(k => k.Key.Split('#')[0] == arquivoScript.ToString()))
                    {
                        string tipo = item.Key.Split('#')[1];
                        string campo = item.Key.Split('#')[2];
                        string valor = item.Value.ToString();

                        if (CampoSempreString(campo))
                        {
                            tipo = "S";
                        }

                        if (valor.Contains("|"))
                        {
                            if (tipo == "S")
                            {
                                valor = "'" + valor.Replace("|", "','") + "'";
                            }
                            else
                            {
                                valor = valor.Replace("|", ",");
                            }

                            script = script.Replace(":" + campo.ToUpper(), valor);
                        }
                        else
                        {
                            script = script.Replace(":" + campo.ToUpper(), RetornaStringPeloTipo(valor));
                        }
                    }

                    try
                    {
                        bd.ExecuteNonQuery(script.Replace("\r", " "));
                        log.Registrar("Script executado com sucesso!" + Environment.NewLine);
                    }
                    catch (Exception Ex)
                    {
                        log.Registrar("Script: " + script + Environment.NewLine, "ERRO");
                        log.Registrar("Erro: " + Ex.Message + Environment.NewLine, "ERRO");
                    }
                }
            }
        }

        private void deletarValores(Log log, Dictionary<string, string> valoresDeletar)
        {
            foreach (var item in valoresDeletar)
            {
                string comando = "delete from {0} where {1} = {2}";

                string tabela = item.Key.Split('.')[0];
                string campo = item.Key.Split('.')[1];
                string valor = item.Value.ToString();

                if (CampoSempreString(campo))
                {
                    comando = String.Format(comando, tabela, campo, "'" + valor + "'");
                }
                else
                {
                    if (isDateTime(valor))
                    {
                        string date = returnDateTime(valor);
                        comando = String.Format(comando, tabela, campo, "TO_DATE('" + date + "','DD/MM/YYYY')");
                    }
                    else if (isInteger(valor))
                    {
                        comando = String.Format(comando, tabela, campo, valor.ToString());
                    }
                    else if (isDouble(valor))
                    {
                        comando = String.Format(comando, tabela, campo, valor.ToString().Replace(',', '.'));
                    }
                    else
                    {
                        comando = String.Format(comando, tabela, campo, "'" + valor + "'");
                    }
                }

                try
                {
                    bd.ExecuteNonQuery(comando);
                    log.Registrar("Comando: " + comando + Environment.NewLine);
                    log.Registrar("Comando executado com sucesso!" + Environment.NewLine);
                }
                catch (Exception Ex)
                {
                    log.Registrar("Comando: " + comando + Environment.NewLine, "ERRO");
                    log.Registrar("Erro: " + Ex.Message + Environment.NewLine, "ERRO");
                }
            }
        }

        private void upsertValores(Log log, Dictionary<string, string> valoresInserir)
        {
            List<string> tabelas = new List<string>();

            foreach (var item in valoresInserir)
            {
                string nomeTabela = item.Key.Split('.')[0].Replace("*", "");
                if (!tabelas.Contains(nomeTabela))
                {
                    tabelas.Add(nomeTabela);
                }
            }

            foreach (var item in tabelas)
            {
                string where = "";

                log.Registrar("# Processando tabela  " + item + Environment.NewLine);
                string comando = "select count(1) from {0} where {1}";

                foreach (var item2 in valoresInserir)
                {
                    if (item2.Key.Contains("*" + item + "."))
                    {
                        string campo = item2.Key.Split('.')[1];
                        string valor = item2.Value;

                        if (CampoSempreString(campo))
                        {
                            valor = "'" + valor + "'";
                        }
                        else
                        {
                            valor = RetornaStringPeloTipo(valor);
                        }

                        if (String.IsNullOrEmpty(where))
                        {
                            where += campo + " = " + valor;
                        }
                        else
                        {
                            where += " and " + campo + " = " + valor;
                        }
                    }
                }

                comando = String.Format(comando, item, where);

                int count = 0;

                try
                {
                    count = Convert.ToInt32(bd.ExecutaComando(comando).Rows[0][0].ToString());
                }
                catch (Exception Ex)
                {
                    log.Registrar("Comando: " + comando + Environment.NewLine, "ERRO");
                    log.Registrar("Erro: " + Ex.Message + Environment.NewLine, "ERRO");
                    continue;
                }

                if (count == 0)
                {
                    // Insert

                    string campos = "";
                    string valores = "";

                    string comandoInsert = "insert into {0} ({1}) values ({2})";

                    foreach (var item2 in valoresInserir)
                    {
                        if (item2.Key.Contains(item + "."))
                        {
                            string campo = item2.Key.Split('.')[1];
                            string valor = item2.Value;

                            if (String.IsNullOrEmpty(campos))
                            {
                                campos += campo;
                            }
                            else
                            {
                                campos += "," + campo;
                            }

                            string valor2 = "";

                            if (CampoSempreString(campo))
                            {
                                valor2 = "'" + valor + "'";
                            }
                            else
                            {
                                valor2 = RetornaStringPeloTipo(valor);
                            }


                            if (String.IsNullOrEmpty(valores))
                            {
                                valores += valor2;
                            }
                            else
                            {
                                valores += "," + valor2;
                            }
                        }
                    }

                    comandoInsert = String.Format(comandoInsert, item, campos, valores);

                    try
                    {
                        bd.ExecuteNonQuery(comandoInsert);
                        log.Registrar("Comando: " + comandoInsert + Environment.NewLine);
                        log.Registrar("Comando executado com sucesso!" + Environment.NewLine);
                    }
                    catch (Exception Ex)
                    {
                        log.Registrar("Comando: " + comandoInsert + Environment.NewLine, "ERRO");
                        log.Registrar("Erro: " + Ex.Message + Environment.NewLine, "ERRO");
                    }
                }
                else
                {
                    // Update

                    string campos = "";

                    string comandoUpdate = "update {0} set {1} where {2}";

                    foreach (var item2 in valoresInserir)
                    {
                        if (item2.Key.Contains(item) && !item2.Key.Contains("*" + item))
                        {
                            string campo = item2.Key.Split('.')[1];
                            string valor = item2.Value;

                            string valor2 = "";

                            if (CampoSempreString(campo))
                            {
                                valor2 = "'" + valor + "'";
                            }
                            else
                            {
                                valor2 = RetornaStringPeloTipo(valor);
                            }

                            if (String.IsNullOrEmpty(campos))
                            {
                                campos += campo + " = " + valor2;
                            }
                            else
                            {
                                campos += ", " + campo + " = " + valor2;
                            }
                        }
                    }

                    comandoUpdate = String.Format(comandoUpdate, item, campos, where);

                    try
                    {
                        bd.ExecuteNonQuery(comandoUpdate);
                        log.Registrar("Comando: " + comandoUpdate + Environment.NewLine);
                        log.Registrar("Comando executado com sucesso!" + Environment.NewLine);
                    }
                    catch (Exception Ex)
                    {
                        log.Registrar("Comando: " + comandoUpdate + Environment.NewLine, "ERRO");
                        log.Registrar("Erro: " + Ex.Message + Environment.NewLine, "ERRO");
                    }
                }

            }
        }

        private string RetornaStringPeloTipo(string valor)
        {
            string valor2 = "";

            if (isDateTime(valor))
            {
                string date = returnDateTime(valor);
                valor2 = "TO_DATE('" + date + "','DD/MM/YYYY')";
            }
            else if (isInteger(valor))
            {
                valor2 = valor.ToString();
            }
            else if (isDouble(valor))
            {
                valor2 = valor.ToString().Replace(',', '.');
            }
            else
            {
                valor2 = "'" + valor.ToString() + "'";
            }

            return valor2;
        }

        private bool isDouble(object Expression)
        {
            double retNum;
            bool isNum = Double.TryParse(Convert.ToString(Expression), System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out retNum);
            return isNum;
        }

        private bool isInteger(object Expression)
        {
            if (Convert.ToString(Expression).Contains(","))
            {
                return false;
            }
            else
            {
                int retNum;
                bool isNum = Int32.TryParse(Convert.ToString(Expression), System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out retNum);
                return isNum;
            }
        }

        private bool isDateTime(object Expression)
        {
            string[] formats = { "dd/MM/yyyy hh:mm:ss", "dd/MM/yyyy" };

            DateTime date;
            bool isDate = DateTime.TryParseExact(Convert.ToString(Expression),
                                                 formats,
                                                 new CultureInfo("pt-BR"),
                                                 DateTimeStyles.None,
                                                 out date);
            return isDate;
        }

        private string returnDateTime(object Expression)
        {
            string[] formats = { "dd/MM/yyyy hh:mm:ss", "dd/MM/yyyy" };

            DateTime date;
            bool isDate = DateTime.TryParseExact(Convert.ToString(Expression),
                                                 formats,
                                                 new CultureInfo("pt-BR"),
                                                 DateTimeStyles.None,
                                                 out date);
            if (!isDate)
            {
                date = DateTime.Now;
            }

            return date.ToString("dd/MM/yyyy", new CultureInfo("pt-BR"));
        }

        private void processarArquivoExcel(Log log, string arquivo,
            ref Dictionary<string, string> valoresDeletar,
            ref Dictionary<string, string> valoresInserir,
            ref Dictionary<string, string> valoresExecutar,
            ref Dictionary<string, string> valoresSaida,
            ref Dictionary<string, string> valoresVerificar,
            ref Dictionary<string, string> valoresScript
            )
        {
            //Create COM Objects. Create a COM object for everything that is referenced
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(arquivo);

            for (int abas = 1; abas <= xlWorkbook.Sheets.Count; abas++)
            {
                log.Registrar("Processando aba " + abas.ToString() + Environment.NewLine);

                Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[abas];
                Excel.Range xlRange = xlWorksheet.UsedRange;

                int rowCount = xlRange.Rows.Count;
                int colCount = xlRange.Columns.Count;

                for (int x = 1; x <= rowCount; x++)
                {
                    for (int y = 1; y <= colCount; y++)
                    {
                        string valorCelula = "";

                        try
                        {
                            valorCelula = xlRange.Cells[x, y].Value2.ToString();
                        }
                        catch (Exception)
                        {
                        }

                        if (valorCelula == "[DELETAR]")
                        {
                            string numTeste = xlRange.Cells[x + 2, y].Value2.ToString();
                            string tabela = xlRange.Cells[x, y + 1].Value2.ToString();

                            for (int resto = y + 1; resto <= colCount; resto++)
                            {
                                string valor = "";
                                string campo = "";

                                try
                                {
                                    valor = xlRange.Cells[x + 2, resto].Value2.ToString();
                                    campo = xlRange.Cells[x + 1, resto].Value2.ToString();
                                }
                                catch (Exception)
                                {
                                }

                                if (!String.IsNullOrEmpty(valor) && !String.IsNullOrEmpty(campo))
                                {
                                    log.Registrar("Deletar " + tabela + "." + campo + ": " + valor + Environment.NewLine);
                                    valoresDeletar.Add(tabela + "." + campo, valor);
                                }
                            }

                            x = x + 3;
                            break;
                        }

                        if (valorCelula == "[SCRIPT]")
                        {
                            string numTeste = xlRange.Cells[x + 2, y].Value2.ToString();
                            string script = xlRange.Cells[x, y + 1].Value2.ToString();
                            bool isDate = false;

                            for (int resto = y + 1; resto <= colCount; resto++)
                            {
                                string valor = "";
                                string campo = "";
                                bool bold = false;

                                try
                                {
                                    valor = xlRange.Cells[x + 2, resto].Value2.ToString();
                                    campo = xlRange.Cells[x + 1, resto].Value2.ToString();

                                    if (campo.StartsWith("DT") || campo.StartsWith("DATA"))
                                    {
                                        double d = double.Parse(valor);
                                        DateTime conv = DateTime.FromOADate(d);
                                        valor = conv.ToString("dd/MM/yyyy", new CultureInfo("pt-BR"));
                                        isDate = true;
                                    }
                                }
                                catch (Exception)
                                {
                                }

                                try
                                {
                                    bold = xlRange.Cells[x + 1, resto].Font.Bold;
                                }
                                catch (Exception)
                                {
                                }

                                if (!String.IsNullOrEmpty(valor) && !String.IsNullOrEmpty(campo))
                                {
                                    if (isDate)
                                    {
                                        log.Registrar("Script " + script + " parâmetro(date) " + campo + ": " + valor + Environment.NewLine);
                                        valoresScript.Add(script + "#D#" + campo, valor);
                                    }
                                    else
                                    {
                                        if (bold)
                                        {
                                            log.Registrar("Script " + script + " parâmetro(string) " + campo + ": " + valor + Environment.NewLine);
                                            valoresScript.Add(script + "#S#" + campo, valor);
                                        }
                                        else
                                        {
                                            log.Registrar("Script " + script + " parâmetro(number) " + campo + ": " + valor + Environment.NewLine);
                                            valoresScript.Add(script + "#N#" + campo, valor);
                                        }
                                    }
                                }
                            }

                            x = x + 3;
                            break;
                        }

                        if (valorCelula == "[INSERIR]")
                        {
                            string numTeste = xlRange.Cells[x + 2, y].Value2.ToString();
                            string tabela = xlRange.Cells[x, y + 1].Value2.ToString();

                            for (int resto = y + 1; resto <= colCount; resto++)
                            {
                                string valor = "";
                                string campo = "";
                                bool bold = false;

                                try
                                {
                                    valor = xlRange.Cells[x + 2, resto].Value2.ToString();
                                    campo = xlRange.Cells[x + 1, resto].Value2.ToString();

                                    if (campo.StartsWith("DT") || campo.StartsWith("DATA"))
                                    {
                                        double d = double.Parse(valor);
                                        DateTime conv = DateTime.FromOADate(d);
                                        valor = conv.ToString("dd/MM/yyyy", new CultureInfo("pt-BR"));
                                    }
                                }
                                catch (Exception)
                                {
                                }

                                try
                                {
                                    bold = xlRange.Cells[x + 1, resto].Font.Bold;
                                }
                                catch (Exception)
                                {
                                }

                                if (!String.IsNullOrEmpty(valor) && !String.IsNullOrEmpty(campo))
                                {
                                    if (bold)
                                    {
                                        log.Registrar("Upsert " + "*" + tabela + "." + campo + ": " + valor + Environment.NewLine);
                                        valoresInserir.Add("*" + tabela + "." + campo, valor);
                                    }
                                    else
                                    {
                                        log.Registrar("Upsert " + tabela + "." + campo + ": " + valor + Environment.NewLine);
                                        valoresInserir.Add(tabela + "." + campo, valor);
                                    }
                                }
                            }

                            x = x + 3;
                            break;
                        }

                        if (valorCelula == "[OBJETOS]")
                        {
                            for (int linha = x + 2; linha <= rowCount; linha++)
                            {
                                string numTeste = "";

                                try
                                {
                                    numTeste = xlRange.Cells[linha, y].Value2.ToString();
                                }
                                catch (Exception)
                                {
                                    continue;
                                }

                                string nomeobjeto = xlRange.Cells[linha, y + 1].Value2.ToString();
                                string nomefuncao = "";
                                try
                                {
                                    nomefuncao = xlRange.Cells[linha, y + 2].Value2.ToString();
                                }
                                catch (Exception)
                                {
                                }

                                if (!String.IsNullOrEmpty(nomefuncao))
                                {
                                    nomeobjeto = nomeobjeto + "." + nomefuncao;
                                }

                                for (int resto = y + 3; resto <= colCount; resto++)
                                {
                                    string valor = "";
                                    string campo = "";

                                    try
                                    {
                                        valor = xlRange.Cells[linha, resto].Value2.ToString();
                                        campo = xlRange.Cells[x + 1, resto].Value2.ToString();
                                    }
                                    catch (Exception)
                                    {
                                    }

                                    if (!String.IsNullOrEmpty(valor) && !String.IsNullOrEmpty(campo))
                                    {
                                        log.Registrar("Teste " + numTeste + " - Execute " + nomeobjeto + " parametro " + campo + ": " + valor + Environment.NewLine);
                                        valoresExecutar.Add(numTeste + "|" + nomeobjeto + "|" + campo, valor);
                                    }
                                }
                            }

                            x = rowCount + 1;
                            break;
                        }

                        if (valorCelula == "[RESULTADO]")
                        {
                            for (int linha = x + 2; linha <= rowCount; linha++)
                            {
                                string numTeste = "";

                                try
                                {
                                    numTeste = xlRange.Cells[linha, y].Value2.ToString();
                                }
                                catch (Exception)
                                {
                                    continue;
                                }

                                for (int resto = y + 1; resto <= colCount; resto++)
                                {
                                    string valor = "";
                                    string campo = "";

                                    try
                                    {
                                        valor = xlRange.Cells[linha, resto].Value2.ToString();
                                        campo = xlRange.Cells[x + 1, resto].Value2.ToString();
                                    }
                                    catch (Exception)
                                    {
                                    }

                                    if (!String.IsNullOrEmpty(valor) && !String.IsNullOrEmpty(campo))
                                    {
                                        log.Registrar("Teste " + numTeste + " - parametro de saída " + campo + ": " + valor + Environment.NewLine);
                                        valoresSaida.Add(numTeste + "|" + campo, valor);
                                    }
                                }
                            }

                            x = rowCount + 1;
                            break;
                        }

                        if (valorCelula == "[BANCO]")
                        {
                            string numTeste = xlRange.Cells[x + 2, y].Value2.ToString();
                            string tabela = xlRange.Cells[x, y + 1].Value2.ToString();

                            for (int resto = y + 1; resto <= colCount; resto++)
                            {
                                string valor = "";
                                string campo = "";
                                bool bold = false;

                                try
                                {
                                    valor = xlRange.Cells[x + 2, resto].Value2.ToString();
                                    campo = xlRange.Cells[x + 1, resto].Value2.ToString();
                                }
                                catch (Exception)
                                {
                                }

                                try
                                {
                                    bold = xlRange.Cells[x + 1, resto].Font.Bold;
                                }
                                catch (Exception)
                                {
                                }

                                if (!String.IsNullOrEmpty(valor) && !String.IsNullOrEmpty(campo))
                                {
                                    if (bold)
                                    {
                                        log.Registrar("Verificar " + "*" + tabela + "." + campo + ": " + valor + Environment.NewLine);
                                        valoresVerificar.Add("*" + tabela + "." + campo, valor);
                                    }
                                    else
                                    {
                                        log.Registrar("Verificar " + tabela + "." + campo + ": " + valor + Environment.NewLine);
                                        valoresVerificar.Add(tabela + "." + campo, valor);
                                    }
                                }
                            }

                            x = x + 3;
                            break;
                        }
                    }
                }

                //cleanup
                GC.Collect();
                GC.WaitForPendingFinalizers();

                //rule of thumb for releasing com objects:
                //  never use two dots, all COM objects must be referenced and released individually
                //  ex: [somthing].[something].[something] is bad

                //release com objects to fully kill excel process from running in the background
                Marshal.ReleaseComObject(xlRange);
                Marshal.ReleaseComObject(xlWorksheet);

            }

            //close and release
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);

            //quit and release
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);
        }

    }
}
