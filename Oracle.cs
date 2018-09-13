using System;
using System.Collections.Generic;
using System.Web;
using System.Data;
using System.IO;
using System.Configuration;
using System.Data.OracleClient;

namespace ExeCasoTeste
{
    public class Oracle
    {
        string connection = "";

        public Oracle(string _stringConn)
        {
            connection = _stringConn;
        }

        private OracleConnection GetConnection()
        {
            //String de Conexão
            return new OracleConnection(connection);
        }

        public Dictionary<string, string> TesteObjetoBanco(string nome, Dictionary<string, string> entrada, Dictionary<string, string> saida)
        {
            Dictionary<string, string> retorno = new Dictionary<string, string>();
            List<OracleParameter> listaRetorno = new List<OracleParameter>();

            try
            {
                OracleCommand cmd = new OracleCommand();
                cmd.Connection = GetConnection();
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandText = nome;

                // parametros de entrada
                foreach (var itemEntrada in entrada)
                {
                    string tipo = itemEntrada.Value.Split('|')[0];
                    string parametro = itemEntrada.Key; 
                    string valor = itemEntrada.Value.Split('|')[1];

                    OracleParameter p = new OracleParameter();

                    p.Direction = ParameterDirection.Input;
                    p.ParameterName = parametro;

                    if (tipo == "D")
                    {
                        p.OracleType = OracleType.DateTime;
                        p.Value = Convert.ToDateTime(valor);
                    }
                    else if (tipo == "F")
                    {
                        p.OracleType = OracleType.Float;
                        p.Value = Convert.ToDouble(valor);
                    }
                    else if (tipo == "I")
                    {
                        p.OracleType = OracleType.Int32;
                        p.Value = Convert.ToInt32(valor);
                    }
                    else if (tipo == "B")
                    {
                        p.OracleType = OracleType.Char;
                        if (valor.ToUpper() == "TRUE")
                        {
                            p.Value = "1";
                        }
                        else
                        {
                            p.Value = "0";
                        }
                    }
                    else
                    {
                        p.OracleType = OracleType.VarChar;
                        p.Size = 500;
                        p.Value = valor;
                    }
                    cmd.Parameters.Add(p);
                }

                foreach (var itemSaida in saida)
                {
                    bool isResult = itemSaida.Value.Split('|')[0].Contains("*");
                    string tipo = itemSaida.Value.Split('|')[0].Replace("*","");
                    string parametro = itemSaida.Key;
                    string valor = itemSaida.Value.Split('|')[1];

                    OracleParameter p = new OracleParameter();

                    if (isResult)
                    {
                        p.Direction = ParameterDirection.ReturnValue;
                    }
                    else
                    {
                        p.Direction = ParameterDirection.Output;
                    }
                    
                    p.ParameterName = parametro;

                    if (tipo == "D")
                    {
                        p.OracleType = OracleType.DateTime;
                    }
                    else if (tipo == "F")
                    {
                        p.OracleType = OracleType.Float;
                    }
                    else if (tipo == "I")
                    {
                        p.OracleType = OracleType.Int32;
                    }
                    else
                    {
                        p.OracleType = OracleType.VarChar;
                        p.Size = 500;
                    }
                    OracleParameter o = cmd.Parameters.Add(p);
                    listaRetorno.Add(o);
                }

                cmd.Connection.Open();

                cmd.ExecuteNonQuery();

                foreach (OracleParameter item in listaRetorno)
                {
                    retorno.Add(item.ParameterName, item.Value.ToString());
                }

                cmd.Connection.Close();
            }
            catch (Exception Ex)
            {
                throw new Exception("Erro ao executar objeto: " + Ex.Message);
            }
            return retorno;
        }

        public string RecuperaCLOB(string comando)
        {
            string ret = "";
            try
            {
                OracleCommand cmd = new OracleCommand(comando);
                cmd.Connection = GetConnection();
                cmd.Connection.Open();
                OracleDataReader reader = cmd.ExecuteReader();
                reader.Read();
                OracleLob CLOB = reader.GetOracleLob(0);
                ret = CLOB.Value.ToString();
                cmd.Connection.Close();
            }
            catch (Exception)
            {
                ret = "";
            }
            return ret;
        }

        public string InsereCLOB(string texto, string comando, string nomeParam)
        {
            try
            {
                OracleCommand cmd = new OracleCommand(comando);
                cmd.Connection = GetConnection();
                cmd.Connection.Open();
                cmd.CommandType = CommandType.Text;
                cmd.Parameters.Add(nomeParam, OracleType.Clob);
                cmd.Parameters[nomeParam].Value = texto;
                cmd.ExecuteNonQuery();
                cmd.Connection.Close();
                return "OK";
            }
            catch (Exception Ex)
            {
                return Ex.Message;
            }
        }

        public string Testa()
        {
            try
            {
                OracleConnection cn = GetConnection();
                cn.Open();
                cn.Close();
                return "OK";
            }
            catch (Exception E)
            {
                return E.Message + Environment.NewLine + E.StackTrace;
            }
        }

        public DataTable ExecutaComando(string comando)
        {
            OracleConnection cn = new OracleConnection();
            OracleCommand dbCommand = cn.CreateCommand();
            DataTable oDt = new DataTable();

            cn = GetConnection();

            dbCommand.CommandText = comando;
            dbCommand.CommandType = CommandType.Text;

            try
            {
                // Seta a conexão no comando
                dbCommand.Connection = cn;

                // Abre a conexão
                cn.Open();

                //Criar um objeto Oracle Data Adapter
                OracleDataAdapter oDa = new OracleDataAdapter(dbCommand);

                //Preenchendo o DataTable
                oDa.Fill(oDt);

                //Resultado da Função
                return oDt;


            }
            catch (Exception ex)
            {
                if (cn.State == ConnectionState.Open)
                {
                    cn.Close();
                }
                dbCommand.Dispose();
                cn.Dispose();
                throw ex;

            }
            finally
            {
                if (cn.State == ConnectionState.Open)
                {
                    cn.Close();
                }
                dbCommand.Dispose();
                cn.Dispose();
            }

        }

        public void ExecuteNonQuery(string comando)
        {
            OracleConnection cn = GetConnection();
            OracleCommand dbCommand = cn.CreateCommand();
            dbCommand.CommandText = comando;
            //dbCommand.CommandType = CommandType.Text;

            try
            {
                // Seta a conexão no comando
                dbCommand.Connection = cn;

                // Abre a conexão
                cn.Open();

                dbCommand.ExecuteNonQuery();


            }
            catch (Exception ex)
            {
                if (cn.State == ConnectionState.Open)
                {
                    cn.Close();
                }
                dbCommand.Dispose();
                cn.Dispose();
                throw ex;

            }
            finally
            {
                if (cn.State == ConnectionState.Open)
                {
                    cn.Close();
                }
                dbCommand.Dispose();
                cn.Dispose();
            }
        }
    }
}