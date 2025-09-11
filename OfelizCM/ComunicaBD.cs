using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Windows.Forms;

namespace OfelizCM
{
    internal class ComunicaBD
    {
        SqlConnection MiConexion = new SqlConnection("Data Source=GALILEU\\PREPARACAO;Initial Catalog=TempoPreparacao;Persist Security Info=True;User ID=SA;Password=preparacao");

        public void ConectarBD()
        {
            if (MiConexion.State == ConnectionState.Closed)
            {
                MiConexion.Open();
            }
        }

        public void DesonectarBD()
        {
            if (MiConexion.State == ConnectionState.Open)
            {
                MiConexion.Close();
            }
        }
        public SqlConnection GetConnection()
        {
            return MiConexion;
        }

        // Método para buscar dados da base de dados e retornar um DataTable
        public DataTable Procurarbd(string Query)
        {
            SqlCommand MiComando = new SqlCommand(Query, MiConexion);
            SqlDataAdapter dataAdapter = new SqlDataAdapter(MiComando);
            DataTable dataTable = new DataTable();

            // Preenchendo o DataTable com os dados da consulta
            dataAdapter.Fill(dataTable);
            return dataTable;
        }

        public List<string> Procurarbdlist(string Query)
        {
            SqlCommand MiComando = new SqlCommand(Query, MiConexion);
            List<string> Result = new List<string>();

            using (SqlDataReader reader = MiComando.ExecuteReader())
            {
                while (reader.Read())
                {
                    for (int i = 0; i < reader.VisibleFieldCount; i++)
                    {
                        Result.Add(reader[i].ToString());
                    }
                }
            }
            return Result;
        }

        public DataTable BuscarRegistros(SqlCommand command)
        {
            SqlDataAdapter dataAdapter = new SqlDataAdapter(command);
            DataTable dataTable = new DataTable();
            dataAdapter.Fill(dataTable);
            return dataTable;
        }
          
    }


    class ComunicaBDprimavera
    {
        SqlConnection MiConexion = new SqlConnection("Data Source=TESLA\\PRIMAVERA;Initial Catalog=PRIOFELIZ;Persist Security Info=True;User ID=CM;Password=OF€l1z201");

        public void ConectarBD()
        {
            MiConexion.Open();
        }
        public void DesonectarBD()
        {
            MiConexion.Close();
        }
        public List<string> Procurarbd(string Query)
        {
            SqlCommand MiComando = new SqlCommand(Query, MiConexion);
            List<string> Result = new List<string>();

            using (SqlDataReader reader = MiComando.ExecuteReader())
            {
                while (reader.Read())
                {
                    for (int i = 0; i < reader.VisibleFieldCount; i++)
                    {
                        Result.Add(reader[i].ToString());
                    }
                }
            }
            return Result;
        }
        public SqlConnection GetConnection()
        {
            return MiConexion;
        }

        public bool dbOperations(List<string> dbOperations)
        {
            bool SUCESSFULL = true;

            using (MiConexion = new SqlConnection("Data Source=TESLA\\PRIMAVERA;Initial Catalog=PRIOFELIZ;Persist Security Info=True;User ID=CM;Password=OF€l1z201"))
            {
                MiConexion.Open();
                SqlTransaction transaction = MiConexion.BeginTransaction();

                foreach (string commandString in dbOperations)
                {
                    SqlCommand cmd = new SqlCommand(commandString, MiConexion, transaction);
                    cmd.ExecuteNonQuery();
                }
                try
                {
                    transaction.Commit();
                }
                catch (Exception EX)
                {
                    transaction.Rollback();
                    SUCESSFULL = false;
                    MessageBox.Show("ERRO A SALVAR DADOS" + Environment.NewLine + EX.Message, "ERRO", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                MiConexion.Close();
            }
            return SUCESSFULL;
        }

    }
}