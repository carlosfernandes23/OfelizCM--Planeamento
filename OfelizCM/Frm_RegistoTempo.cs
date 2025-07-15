using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OfelizCM
{
    public partial class Frm_RegistoTempo : Form
    {
        public Frm_RegistoTempo()
        {
            InitializeComponent();
            DataGridViewRegistoTempo.ClearSelection();
            DateTimePickerInicio.Value = DateTime.Now;
            DateTimePickerApioaaObra.Value = DateTime.Now;
        }

        private void Frm_RegistoTempo_Load(object sender, EventArgs e)
        {

            ConfirmarComunicacaoBD();
            CarregarPrioridadesNaComboBox();
            CarregarPreparadoresNaComboBox();
            VerificarUsuario();

            foreach (DataGridViewColumn column in DataGridViewRegistoTempo.Columns)
            {
                column.ReadOnly = true;
            }

            if (DataGridViewRegistoTempo.Columns.Contains("Hora Inicial"))
            {
                DataGridViewRegistoTempo.Columns["Hora Inicial"].ReadOnly = false;
            }

            if (DataGridViewRegistoTempo.Columns.Contains("Hora Final"))
            {
                DataGridViewRegistoTempo.Columns["Hora Final"].ReadOnly = false;
            }

            if (DataGridViewRegistoTempo.Columns.Contains("ObservaçõesPreparador"))
            {
                DataGridViewRegistoTempo.Columns["ObservaçõesPreparador"].ReadOnly = false;
            }

            DataGridViewRegistoTempo.CellFormatting += DataGridViewRegistoTempo_CellFormatting;

            


        }
        private void DataGridViewRegistoTempo_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
        }

        private void DataGridViewRegistoTempo_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {

            if (DataGridViewRegistoTempo.Columns[e.ColumnIndex].Name == "ObservaçõesPreparador")
            {
                if (e.Value != null && e.Value.ToString() == "0")
                {
                    e.Value = "";
                }
            }
            DataGridViewRegistoTempo.Columns["ObservaçõesPreparador"].HeaderText = "Observações do Preparador";


        }

        public void CarregarPrioridadesNaComboBox()
        {
            ComunicaBD BD = new ComunicaBD();
            BD.ConectarBD();
            string query = "SELECT Prioridade FROM dbo.Prioridades";
            List<string> list = BD.Procurarbdlist(query);
            BD.DesonectarBD();

            ComboBoxPrioAdd.Items.Clear();
            if (list.Count > 0)
            {
                foreach (string nome in list)
                {
                    ComboBoxPrioAdd.Items.Add(nome);
                }
            }
            else
            {
                MessageBox.Show("Nenhum preparador encontrado na base de dados.");

            }
        }
        //private void ComunicaBDparaTabela()
        //{
        //    string nomePreparador = Environment.UserName;
        //    string[] partes = nomePreparador.Split('.');
        //    List<string> partesComMaiusculas = new List<string>();

        //    foreach (string parte in partes)
        //    {
        //        if (!string.IsNullOrEmpty(parte))
        //        {
        //            string parteFormatada = char.ToUpper(parte[0]) + parte.Substring(1).ToLower();
        //            partesComMaiusculas.Add(parteFormatada);
        //        }
        //    }
        //    string nomeFormatado = string.Join(" ", partesComMaiusculas);

        //    ComunicaBD BD = new ComunicaBD();

        //    try
        //    {
        //        BD.ConectarBD();  

        //        string query = "SELECT ID, [Numero da Obra], [Nome da Obra], Tarefa, Preparador, [Hora Inicial], [Hora Final], [Data da Tarefa], [Qtd de Hora], Prioridade, ObservaçõesPreparador " +
        //                       "FROM dbo.RegistoTempo " +
        //                       "WHERE Preparador = '" + nomeFormatado + "'";  

        //        DataTable dataTable = BD.Procurarbd(query);  

        //        DataGridViewRegistoTempo.DataSource = dataTable;
        //        DataGridViewRegistoTempo.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
        //        DataGridViewRegistoTempo.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
        //        DataGridViewRegistoTempo.Columns["ID"].Visible = false;
        //        DataGridViewRegistoTempo.Columns["Preparador"].Visible = false;
        //        DataGridViewRegistoTempo.ClearSelection();
        //        DataGridViewRegistoTempo.AutoResizeColumns();
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show("Erro ao conectar à base de dados: " + ex.Message);
        //    }
        //    finally
        //    {
        //        BD.DesonectarBD();  
        //    }
        //}


        //private void ComunicaBDparaTabelafiltrocomData()
        //{
        //    ComunicaBD comunicaBD = new ComunicaBD();

        //    try
        //    {
        //        comunicaBD.ConectarBD();

        //        string nomePreparador = Environment.UserName;
        //        string[] partes = nomePreparador.Split('.');
        //        List<string> partesComMaiusculas = new List<string>();

        //        foreach (string parte in partes)
        //        {
        //            if (!string.IsNullOrEmpty(parte))
        //            {
        //                string parteFormatada = char.ToUpper(parte[0]) + parte.Substring(1).ToLower();
        //                partesComMaiusculas.Add(parteFormatada);
        //            }
        //        }

        //        string nomeFormatado = string.Join(" ", partesComMaiusculas);

        //        string Obra = TextBoxNObra.Text;

        //        string Prioridades = null;

        //        if (ComboBoxPrioAdd.SelectedItem != null)
        //        {
        //            Prioridades = ComboBoxPrioAdd.SelectedItem.ToString();
        //        }

        //        string query = "SELECT Id, [Numero da Obra], [Nome da Obra], Tarefa, Preparador, [Hora Inicial], [Hora Final], [Data da Tarefa], [Qtd de Hora], Prioridade, ObservaçõesPreparador " +
        //                       "FROM dbo.RegistoTempo WHERE 1=1";

        //        query += " AND Preparador = @NomePreparador";

        //        if (DateTimePickerInicio.Value != DateTimePickerInicio.MinDate)
        //        {
        //            query += " AND TRY_CONVERT(DATETIME, [Data da Tarefa], 103) = @DataInicio";
        //        }

        //        if (!string.IsNullOrEmpty(TextBoxNObra.Text))
        //        {
        //            query += " AND [Numero da Obra] = @NumeroObra";
        //        }

        //        if (!string.IsNullOrEmpty(Prioridades))
        //        {
        //            query += " AND Prioridade = @Prioridade";
        //        }

        //        DataTable dataTable = new DataTable();

        //        using (var command = new SqlCommand(query, comunicaBD.GetConnection()))
        //        {
        //            command.Parameters.AddWithValue("@NomePreparador", nomeFormatado);

        //            if (DateTimePickerInicio.Value != DateTimePickerInicio.MinDate)
        //            {
        //                command.Parameters.AddWithValue("@DataInicio", DateTimePickerInicio.Value.Date);
        //            }

        //            if (!string.IsNullOrEmpty(TextBoxNObra.Text))
        //            {
        //                command.Parameters.AddWithValue("@NumeroObra", Obra);
        //            }

        //            if (!string.IsNullOrEmpty(Prioridades))
        //            {
        //                command.Parameters.AddWithValue("@Prioridade", Prioridades);
        //            }

        //            using (var adapter = new SqlDataAdapter(command))
        //            {
        //                adapter.Fill(dataTable);
        //            }
        //        }

        //        DataGridViewRegistoTempo.DataSource = dataTable;
        //        DataGridViewRegistoTempo.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
        //        DataGridViewRegistoTempo.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
        //        DataGridViewRegistoTempo.Columns["Id"].Visible = false;
        //        DataGridViewRegistoTempo.Columns["Preparador"].Visible = false;
        //        DataGridViewRegistoTempo.ClearSelection();
        //        DataGridViewRegistoTempo.AutoResizeColumns();
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show("Erro ao conectar à base de dados: " + ex.Message);
        //    }
        //    finally
        //    {
        //        comunicaBD.DesonectarBD();
        //    }
        //}


        private void ConfirmarComunicacaoBD()
        {
            string nomePreparador = Environment.UserName;
            string nomeUsuario2 = Properties.Settings.Default.NomeUsuario;

            if (nomePreparador.Equals("helder.silva", StringComparison.OrdinalIgnoreCase) ||
                nomeUsuario2.Equals("helder.silva", StringComparison.OrdinalIgnoreCase))
            {
                ComunicaBDparaTabelaHelder1semanas();
            }
            else
            {
                ComunicaBDparaTabelaTodos();
            }
        }

        private void ConfirmarComunicacaoBDSemana1()
        {
            string nomePreparador = Environment.UserName;
            string nomeUsuario2 = Properties.Settings.Default.NomeUsuario;

            if (nomePreparador.Equals("helder.silva", StringComparison.OrdinalIgnoreCase) ||
                nomeUsuario2.Equals("helder.silva", StringComparison.OrdinalIgnoreCase))
            {
                ComunicaBDparaTabelaHelder1semanas();
            }
            else
            {
                ComunicaBDparaTabelaTodos1semanas();
            }
        }

        private void ConfirmarComunicacaoBDSemana2()
        {
            string nomePreparador = Environment.UserName;
            string nomeUsuario2 = Properties.Settings.Default.NomeUsuario;

            if (nomePreparador.Equals("helder.silva", StringComparison.OrdinalIgnoreCase) ||
                nomeUsuario2.Equals("helder.silva", StringComparison.OrdinalIgnoreCase))
            {
                ComunicaBDparaTabelaHelder2semanas();
            }
            else
            {
                ComunicaBDparaTabelaTodos2semanas();
            }
        }

        private void ConfirmarComunicacaoBDSemana3()
        {
            string nomePreparador = Environment.UserName;
            string nomeUsuario2 = Properties.Settings.Default.NomeUsuario;

            if (nomePreparador.Equals("helder.silva", StringComparison.OrdinalIgnoreCase) ||
                nomeUsuario2.Equals("helder.silva", StringComparison.OrdinalIgnoreCase))
            {
                ComunicaBDparaTabelaHelder3semanas();
            }
            else
            {
                ComunicaBDparaTabelaTodos3semanas();
            }
        }

        private void ConfirmarComunicacaoBDSemana4()
        {
            string nomePreparador = Environment.UserName;
            string nomeUsuario2 = Properties.Settings.Default.NomeUsuario;

            if (nomePreparador.Equals("helder.silva", StringComparison.OrdinalIgnoreCase) ||
                nomeUsuario2.Equals("helder.silva", StringComparison.OrdinalIgnoreCase))
            {
                ComunicaBDparaTabelaHelder4semanas();
            }
            else
            {
                ComunicaBDparaTabelaTodos4semanas();
            }
        }

        private void ComunicaBDparaTabelaTodos()
        {
            string nomePreparador = Environment.UserName;
            string[] partes = nomePreparador.Split('.');
            List<string> partesComMaiusculas = new List<string>();

            foreach (string parte in partes)
            {
                if (!string.IsNullOrEmpty(parte))
                {
                    string parteFormatada = char.ToUpper(parte[0]) + parte.Substring(1).ToLower();
                    partesComMaiusculas.Add(parteFormatada);
                }
            }
            string nomeFormatado = string.Join(" ", partesComMaiusculas);

            ComunicaBD BD = new ComunicaBD();

            try
            {
                BD.ConectarBD();

                string query = "SELECT ID, [Numero da Obra], [Nome da Obra], Tarefa, Preparador, [Hora Inicial], [Hora Final], [Data da Tarefa], [Qtd de Hora], Prioridade, ObservaçõesPreparador " +
                               "FROM dbo.RegistoTempo " +
                               "WHERE Preparador = '" + nomeFormatado + "' " +
                               "AND CONVERT(VARCHAR, [Data da Tarefa], 103) = CONVERT(VARCHAR, GETDATE(), 103) " +  
                               "ORDER BY [Data da Tarefa] ASC";


                DataTable dataTable = BD.Procurarbd(query);

                DataGridViewRegistoTempo.DataSource = dataTable;
                DataGridViewRegistoTempo.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
                DataGridViewRegistoTempo.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
                DataGridViewRegistoTempo.Columns["ID"].Visible = false;
                DataGridViewRegistoTempo.Columns["Preparador"].Visible = false;
                DataGridViewRegistoTempo.ClearSelection();
                DataGridViewRegistoTempo.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                MessageBox.Show(" Todos - Erro ao conectar à base de dados: " + ex.Message);
            }
            finally
            {
                BD.DesonectarBD();
            }
        }

        private void ComunicaBDparaTabelaHelder1semanas()
        {

            ComunicaBD BD = new ComunicaBD();

            try
            {
                BD.ConectarBD();


                string query = "SELECT ID, Preparador, [Numero da Obra], [Nome da Obra], Tarefa, [Hora Inicial], [Hora Final], [Data da Tarefa], [Qtd de Hora], Prioridade, ObservaçõesPreparador " +
                               "FROM dbo.RegistoTempo " +
                               "WHERE DATEDIFF(day, TRY_CONVERT(DATE, [Data da Tarefa], 103), TRY_CONVERT(DATE, GETDATE(), 103)) <= 7 " +
                               "ORDER BY [Data da Tarefa] ASC";

                DataTable dataTable = BD.Procurarbd(query);

                DataGridViewRegistoTempo.DataSource = dataTable;
                DataGridViewRegistoTempo.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
                DataGridViewRegistoTempo.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
                DataGridViewRegistoTempo.Columns["ID"].Visible = false;
                DataGridViewRegistoTempo.ClearSelection();
                DataGridViewRegistoTempo.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                MessageBox.Show("HELDER 1 - Erro ao conectar à base de dados: " + ex.Message);
            }
            finally
            {
                BD.DesonectarBD();
            }
        }

        private void ComunicaBDparaTabelaTodos1semanas()
        {
            string nomePreparador = Environment.UserName;
            string[] partes = nomePreparador.Split('.');
            List<string> partesComMaiusculas = new List<string>();

            foreach (string parte in partes)
            {
                if (!string.IsNullOrEmpty(parte))
                {
                    string parteFormatada = char.ToUpper(parte[0]) + parte.Substring(1).ToLower();
                    partesComMaiusculas.Add(parteFormatada);
                }
            }
            string nomeFormatado = string.Join(" ", partesComMaiusculas);

            ComunicaBD BD = new ComunicaBD();

            try
            {
                BD.ConectarBD();

                string query = "SELECT ID, [Numero da Obra], [Nome da Obra], Tarefa, Preparador, [Hora Inicial], [Hora Final], [Data da Tarefa], [Qtd de Hora], Prioridade, ObservaçõesPreparador " +
                               "FROM dbo.RegistoTempo " +
                               "WHERE Preparador = '" + nomeFormatado + "' " +
                               "AND DATEDIFF(day, TRY_CONVERT(DATE, [Data da Tarefa], 103), TRY_CONVERT(DATE, GETDATE(), 103)) <= 7" +
                               "ORDER BY [Data da Tarefa] ASC";



                DataTable dataTable = BD.Procurarbd(query);

                DataGridViewRegistoTempo.DataSource = dataTable;
                DataGridViewRegistoTempo.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
                DataGridViewRegistoTempo.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
                DataGridViewRegistoTempo.Columns["ID"].Visible = false;
                DataGridViewRegistoTempo.Columns["Preparador"].Visible = false;
                DataGridViewRegistoTempo.ClearSelection();
                DataGridViewRegistoTempo.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Todos 1 - Erro ao conectar à base de dados: " + ex.Message);
            }
            finally
            {
                BD.DesonectarBD();
            }
        }

        private void ComunicaBDparaTabelaTodos2semanas()
        {
            string nomePreparador = Environment.UserName;
            string[] partes = nomePreparador.Split('.');
            List<string> partesComMaiusculas = new List<string>();

            foreach (string parte in partes)
            {
                if (!string.IsNullOrEmpty(parte))
                {
                    string parteFormatada = char.ToUpper(parte[0]) + parte.Substring(1).ToLower();
                    partesComMaiusculas.Add(parteFormatada);
                }
            }
            string nomeFormatado = string.Join(" ", partesComMaiusculas);

            ComunicaBD BD = new ComunicaBD();

            try
            {
                BD.ConectarBD();

                string query = "SELECT ID, [Numero da Obra], [Nome da Obra], Tarefa, Preparador, [Hora Inicial], [Hora Final], [Data da Tarefa], [Qtd de Hora], Prioridade, ObservaçõesPreparador " +
                               "FROM dbo.RegistoTempo " +
                               "WHERE Preparador = '" + nomeFormatado + "' " +
                               "AND DATEDIFF(day, TRY_CONVERT(DATE, [Data da Tarefa], 103), TRY_CONVERT(DATE, GETDATE(), 103)) <= 14" +
                               "ORDER BY [Data da Tarefa] ASC";



                DataTable dataTable = BD.Procurarbd(query);

                DataGridViewRegistoTempo.DataSource = dataTable;
                DataGridViewRegistoTempo.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
                DataGridViewRegistoTempo.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
                DataGridViewRegistoTempo.Columns["ID"].Visible = false;
                DataGridViewRegistoTempo.Columns["Preparador"].Visible = false;
                DataGridViewRegistoTempo.ClearSelection();
                DataGridViewRegistoTempo.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Todos 2 - Erro ao conectar à base de dados: " + ex.Message);
            }
            finally
            {
                BD.DesonectarBD();
            }
        }

        private void ComunicaBDparaTabelaHelder2semanas()
        {

            ComunicaBD BD = new ComunicaBD();

            try
            {
                BD.ConectarBD();

                string query = "SELECT ID, Preparador, [Numero da Obra], [Nome da Obra], Tarefa, [Hora Inicial], [Hora Final], [Data da Tarefa], [Qtd de Hora], Prioridade, ObservaçõesPreparador " +
                               "FROM dbo.RegistoTempo " +
                               "WHERE DATEDIFF(day, TRY_CONVERT(DATE, [Data da Tarefa], 103), TRY_CONVERT(DATE, GETDATE(), 103)) <= 14" +
                               "ORDER BY [Data da Tarefa] ASC";



                DataTable dataTable = BD.Procurarbd(query);

                DataGridViewRegistoTempo.DataSource = dataTable;
                DataGridViewRegistoTempo.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
                DataGridViewRegistoTempo.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
                DataGridViewRegistoTempo.Columns["ID"].Visible = false;
                DataGridViewRegistoTempo.ClearSelection();
                DataGridViewRegistoTempo.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                MessageBox.Show("HELDER 2 - Erro ao conectar à base de dados: " + ex.Message);
            }
            finally
            {
                BD.DesonectarBD();
            }
        }

        private void ComunicaBDparaTabelaTodos3semanas()
        {
            string nomePreparador = Environment.UserName;
            string[] partes = nomePreparador.Split('.');
            List<string> partesComMaiusculas = new List<string>();

            foreach (string parte in partes)
            {
                if (!string.IsNullOrEmpty(parte))
                {
                    string parteFormatada = char.ToUpper(parte[0]) + parte.Substring(1).ToLower();
                    partesComMaiusculas.Add(parteFormatada);
                }
            }
            string nomeFormatado = string.Join(" ", partesComMaiusculas);

            ComunicaBD BD = new ComunicaBD();

            try
            {
                BD.ConectarBD();

                string query = "SELECT ID, [Numero da Obra], [Nome da Obra], Tarefa, Preparador, [Hora Inicial], [Hora Final], [Data da Tarefa], [Qtd de Hora], Prioridade, ObservaçõesPreparador " +
                               "FROM dbo.RegistoTempo " +
                               "WHERE Preparador = '" + nomeFormatado + "' " +
                               "AND DATEDIFF(day, TRY_CONVERT(DATE, [Data da Tarefa], 103), TRY_CONVERT(DATE, GETDATE(), 103)) <= 21" +
                               "ORDER BY [Data da Tarefa] ASC";



                DataTable dataTable = BD.Procurarbd(query);

                DataGridViewRegistoTempo.DataSource = dataTable;
                DataGridViewRegistoTempo.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
                DataGridViewRegistoTempo.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
                DataGridViewRegistoTempo.Columns["ID"].Visible = false;
                DataGridViewRegistoTempo.Columns["Preparador"].Visible = false;
                DataGridViewRegistoTempo.ClearSelection();
                DataGridViewRegistoTempo.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Todos 3 - Erro ao conectar à base de dados: " + ex.Message);
            }
            finally
            {
                BD.DesonectarBD();
            }
        }

        private void ComunicaBDparaTabelaHelder3semanas()
        {

            ComunicaBD BD = new ComunicaBD();

            try
            {
                BD.ConectarBD();

                string query = "SELECT ID, Preparador, [Numero da Obra], [Nome da Obra], Tarefa, [Hora Inicial], [Hora Final], [Data da Tarefa], [Qtd de Hora], Prioridade, ObservaçõesPreparador " +
                               "FROM dbo.RegistoTempo " +
                               "WHERE DATEDIFF(day, TRY_CONVERT(DATE, [Data da Tarefa], 103), TRY_CONVERT(DATE, GETDATE(), 103)) <= 21" +
                               "ORDER BY [Data da Tarefa] ASC";



                DataTable dataTable = BD.Procurarbd(query);

                DataGridViewRegistoTempo.DataSource = dataTable;
                DataGridViewRegistoTempo.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
                DataGridViewRegistoTempo.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
                DataGridViewRegistoTempo.Columns["ID"].Visible = false;
                DataGridViewRegistoTempo.ClearSelection();
                DataGridViewRegistoTempo.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                MessageBox.Show("HELDER 3 - Erro ao conectar à base de dados: " + ex.Message);
            }
            finally
            {
                BD.DesonectarBD();
            }
        }

        private void ComunicaBDparaTabelaTodos4semanas()
        {
            string nomePreparador = Environment.UserName;
            string[] partes = nomePreparador.Split('.');
            List<string> partesComMaiusculas = new List<string>();

            foreach (string parte in partes)
            {
                if (!string.IsNullOrEmpty(parte))
                {
                    string parteFormatada = char.ToUpper(parte[0]) + parte.Substring(1).ToLower();
                    partesComMaiusculas.Add(parteFormatada);
                }
            }
            string nomeFormatado = string.Join(" ", partesComMaiusculas);

            ComunicaBD BD = new ComunicaBD();

            try
            {
                BD.ConectarBD();

                string query = "SELECT ID, [Numero da Obra], [Nome da Obra], Tarefa, Preparador, [Hora Inicial], [Hora Final], [Data da Tarefa], [Qtd de Hora], Prioridade, ObservaçõesPreparador " +
                               "FROM dbo.RegistoTempo " +
                               "WHERE Preparador = '" + nomeFormatado + "' " +
                               "AND DATEDIFF(day, TRY_CONVERT(DATE, [Data da Tarefa], 103), TRY_CONVERT(DATE, GETDATE(), 103)) <= 31" +
                               "ORDER BY [Data da Tarefa] ASC";



                DataTable dataTable = BD.Procurarbd(query);

                DataGridViewRegistoTempo.DataSource = dataTable;
                DataGridViewRegistoTempo.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
                DataGridViewRegistoTempo.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
                DataGridViewRegistoTempo.Columns["ID"].Visible = false;
                DataGridViewRegistoTempo.Columns["Preparador"].Visible = false;
                DataGridViewRegistoTempo.ClearSelection();
                DataGridViewRegistoTempo.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Todos 4 - Erro ao conectar à base de dados: " + ex.Message);
            }
            finally
            {
                BD.DesonectarBD();
            }
        }

        private void ComunicaBDparaTabelaHelder4semanas()
        {

            ComunicaBD BD = new ComunicaBD();

            try
            {
                BD.ConectarBD();

                string query = "SELECT ID, Preparador, [Numero da Obra], [Nome da Obra], Tarefa, [Hora Inicial], [Hora Final], [Data da Tarefa], [Qtd de Hora], Prioridade, ObservaçõesPreparador " +
                               "FROM dbo.RegistoTempo " +
                               "WHERE DATEDIFF(day, TRY_CONVERT(DATE, [Data da Tarefa], 103), TRY_CONVERT(DATE, GETDATE(), 103)) <= 31" +
                               "ORDER BY [Data da Tarefa] ASC";



                DataTable dataTable = BD.Procurarbd(query);

                DataGridViewRegistoTempo.DataSource = dataTable;
                DataGridViewRegistoTempo.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
                DataGridViewRegistoTempo.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
                DataGridViewRegistoTempo.Columns["ID"].Visible = false;
                DataGridViewRegistoTempo.ClearSelection();
                DataGridViewRegistoTempo.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                MessageBox.Show("HELDER 4 - Erro ao conectar à base de dados: " + ex.Message);
            }
            finally
            {
                BD.DesonectarBD();
            }
        }

        private void ComunicaBDparaTabelafiltrocomData()
        {
            ComunicaBD comunicaBD = new ComunicaBD();

            try
            {
                comunicaBD.ConectarBD();

                string nomePreparador = string.Empty;

                if (ComboBoxPreparadorAdd.SelectedItem != null && !string.IsNullOrEmpty(ComboBoxPreparadorAdd.SelectedItem.ToString()))
                {
                    nomePreparador = ComboBoxPreparadorAdd.SelectedItem.ToString();
                }
                else
                {
                    nomePreparador = Environment.UserName;
                    string[] partes = nomePreparador.Split('.');
                    List<string> partesComMaiusculas = new List<string>();

                    foreach (string parte in partes)
                    {
                        if (!string.IsNullOrEmpty(parte))
                        {
                            string parteFormatada = char.ToUpper(parte[0]) + parte.Substring(1).ToLower();
                            partesComMaiusculas.Add(parteFormatada);
                        }
                    }

                    nomePreparador = string.Join(" ", partesComMaiusculas);
                }

                string Obra = TextBoxNObra.Text;

                string Prioridades = null;

                if (ComboBoxPrioAdd.SelectedItem != null)
                {
                    Prioridades = ComboBoxPrioAdd.SelectedItem.ToString();
                }

                string query = "SELECT Id, [Numero da Obra], [Nome da Obra], Tarefa, Preparador, [Hora Inicial], [Hora Final], [Data da Tarefa], [Qtd de Hora], Prioridade, ObservaçõesPreparador " +
                                "FROM dbo.RegistoTempo WHERE 1=1";

                query += " AND Preparador = @NomePreparador";

                if (DateTimePickerInicio.Value != DateTimePickerInicio.MinDate)
                {
                    query += " AND TRY_CONVERT(DATETIME, [Data da Tarefa], 103) = @DataInicio";
                }

                if (!string.IsNullOrEmpty(TextBoxNObra.Text))
                {
                    query += " AND [Numero da Obra] = @NumeroObra";
                }

                if (!string.IsNullOrEmpty(Prioridades))
                {
                    query += " AND Prioridade = @Prioridade";
                }

                DataTable dataTable = new DataTable();

                using (var command = new SqlCommand(query, comunicaBD.GetConnection()))
                {
                    command.Parameters.AddWithValue("@NomePreparador", nomePreparador);

                    if (DateTimePickerInicio.Value != DateTimePickerInicio.MinDate)
                    {
                        command.Parameters.AddWithValue("@DataInicio", DateTimePickerInicio.Value.Date);
                    }

                    if (!string.IsNullOrEmpty(TextBoxNObra.Text))
                    {
                        command.Parameters.AddWithValue("@NumeroObra", Obra);
                    }

                    if (!string.IsNullOrEmpty(Prioridades))
                    {
                        command.Parameters.AddWithValue("@Prioridade", Prioridades);
                    }

                    using (var adapter = new SqlDataAdapter(command))
                    {
                        adapter.Fill(dataTable);
                    }
                }

                DataGridViewRegistoTempo.DataSource = dataTable;
                DataGridViewRegistoTempo.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
                DataGridViewRegistoTempo.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
                DataGridViewRegistoTempo.Columns["Id"].Visible = false;
                DataGridViewRegistoTempo.Columns["Preparador"].Visible = false;
                DataGridViewRegistoTempo.ClearSelection();
                DataGridViewRegistoTempo.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erro ao conectar à base de dados: " + ex.Message);
            }
            finally
            {
                comunicaBD.DesonectarBD();
            }
        }


        //private void ComunicaBDparaTabelafiltrosemData()
        //{
        //    ComunicaBD comunicaBD = new ComunicaBD();

        //    try
        //    {
        //        comunicaBD.ConectarBD();

        //        string nomePreparador = Environment.UserName;
        //        string[] partes = nomePreparador.Split('.');
        //        List<string> partesComMaiusculas = new List<string>();

        //        foreach (string parte in partes)
        //        {
        //            if (!string.IsNullOrEmpty(parte))
        //            {
        //                string parteFormatada = char.ToUpper(parte[0]) + parte.Substring(1).ToLower();
        //                partesComMaiusculas.Add(parteFormatada);
        //            }
        //        }

        //        string nomeFormatado = string.Join(" ", partesComMaiusculas);

        //        string Obra = TextBoxNObra.Text;

        //        string Prioridades = null;

        //        if (ComboBoxPrioAdd.SelectedItem != null)
        //        {
        //            Prioridades = ComboBoxPrioAdd.SelectedItem.ToString();
        //        }

        //        string query = "SELECT Id, [Numero da Obra], [Nome da Obra], Tarefa, Preparador, [Hora Inicial], [Hora Final], [Data da Tarefa], [Qtd de Hora], Prioridade, ObservaçõesPreparador " +
        //                       "FROM dbo.RegistoTempo WHERE 1=1";

        //        query += " AND Preparador = @NomePreparador";


        //        if (!string.IsNullOrEmpty(TextBoxNObra.Text))
        //        {
        //            query += " AND [Numero da Obra] = @NumeroObra";
        //        }

        //        if (!string.IsNullOrEmpty(Prioridades))
        //        {
        //            query += " AND Prioridade = @Prioridade";
        //        }

        //        DataTable dataTable = new DataTable();

        //        using (var command = new SqlCommand(query, comunicaBD.GetConnection()))
        //        {
        //            command.Parameters.AddWithValue("@NomePreparador", nomeFormatado);


        //            if (!string.IsNullOrEmpty(TextBoxNObra.Text))
        //            {
        //                command.Parameters.AddWithValue("@NumeroObra", Obra);
        //            }

        //            if (!string.IsNullOrEmpty(Prioridades))
        //            {
        //                command.Parameters.AddWithValue("@Prioridade", Prioridades);
        //            }

        //            using (var adapter = new SqlDataAdapter(command))
        //            {
        //                adapter.Fill(dataTable);
        //            }
        //        }

        //        DataGridViewRegistoTempo.DataSource = dataTable;
        //        DataGridViewRegistoTempo.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
        //        DataGridViewRegistoTempo.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
        //        DataGridViewRegistoTempo.Columns["Id"].Visible = false;
        //        DataGridViewRegistoTempo.Columns["Preparador"].Visible = false;
        //        DataGridViewRegistoTempo.ClearSelection();
        //        DataGridViewRegistoTempo.AutoResizeColumns();
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show("Erro ao conectar à base de dados: " + ex.Message);
        //    }
        //    finally
        //    {
        //        comunicaBD.DesonectarBD();
        //    }
        //}      

        private void ComunicaBDparaTabelafiltrosemData()
        {
            ComunicaBD comunicaBD = new ComunicaBD();

            try
            {
                comunicaBD.ConectarBD();

                string nomePreparador = string.Empty;

                if (ComboBoxPreparadorAdd.SelectedItem != null && !string.IsNullOrEmpty(ComboBoxPreparadorAdd.SelectedItem.ToString()))
                {
                    nomePreparador = ComboBoxPreparadorAdd.SelectedItem.ToString();
                }
                else
                {
                    nomePreparador = Environment.UserName;
                    string[] partes = nomePreparador.Split('.');
                    List<string> partesComMaiusculas = new List<string>();

                    foreach (string parte in partes)
                    {
                        if (!string.IsNullOrEmpty(parte))
                        {
                            string parteFormatada = char.ToUpper(parte[0]) + parte.Substring(1).ToLower();
                            partesComMaiusculas.Add(parteFormatada);
                        }
                    }

                    nomePreparador = string.Join(" ", partesComMaiusculas);
                }

                string Obra = TextBoxNObra.Text;

                string Prioridades = null;

                if (ComboBoxPrioAdd.SelectedItem != null)
                {
                    Prioridades = ComboBoxPrioAdd.SelectedItem.ToString();
                }

                string query = "SELECT Id, [Numero da Obra], [Nome da Obra], Tarefa, Preparador, [Hora Inicial], [Hora Final], [Data da Tarefa], [Qtd de Hora], Prioridade, ObservaçõesPreparador " +
                               "FROM dbo.RegistoTempo WHERE 1=1";

                query += " AND Preparador = @NomePreparador";

                if (!string.IsNullOrEmpty(TextBoxNObra.Text))
                {
                    query += " AND [Numero da Obra] = @NumeroObra";
                }

                if (!string.IsNullOrEmpty(Prioridades))
                {
                    query += " AND Prioridade = @Prioridade";
                }

                DataTable dataTable = new DataTable();

                using (var command = new SqlCommand(query, comunicaBD.GetConnection()))
                {
                    command.Parameters.AddWithValue("@NomePreparador", nomePreparador);  

                    if (!string.IsNullOrEmpty(TextBoxNObra.Text))
                    {
                        command.Parameters.AddWithValue("@NumeroObra", Obra);
                    }

                    if (!string.IsNullOrEmpty(Prioridades))
                    {
                        command.Parameters.AddWithValue("@Prioridade", Prioridades);
                    }

                    using (var adapter = new SqlDataAdapter(command))
                    {
                        adapter.Fill(dataTable);
                    }
                }

                DataGridViewRegistoTempo.DataSource = dataTable;
                DataGridViewRegistoTempo.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
                DataGridViewRegistoTempo.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
                DataGridViewRegistoTempo.Columns["Id"].Visible = false;
                DataGridViewRegistoTempo.Columns["Preparador"].Visible = false;
                DataGridViewRegistoTempo.ClearSelection();
                DataGridViewRegistoTempo.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erro ao conectar à base de dados: " + ex.Message);
            }
            finally
            {
                comunicaBD.DesonectarBD();
            }
        }


        private void guna2Button1_Click(object sender, EventArgs e)
        {
            ComunicaBDparaTabelafiltrocomData();
        }

        private void ButtonConfirmarTarefa_Click(object sender, EventArgs e)
        {
            ComunicaBDparaTabelafiltrosemData();
        }

      
        private void ButtonAtualizarDados_Click(object sender, EventArgs e)
        {
            CarregarNomeObraPorCaminho();
            EnviarHoraInicioParaBaseDeDados();
            ConfirmarComunicacaoBD();

        }

        private void EnviarHoraInicioParaBaseDeDados()
        {
            int horaInicio = (int)NumericUpDownHInicio.Value;
            int minutoInicio = (int)NumericUpDownMInicio.Value;

            string horaFormatada = horaInicio.ToString("D2");
            string minutoFormatado = minutoInicio.ToString("D2");
            string HoraInicio = $"{horaFormatada}:{minutoFormatado}:00";

            int horafinal = (int)NumericUpDownHFinal.Value;
            int minutofinal = (int)NumericUpDownMFinal.Value;

            string horaFinalFormatada = horafinal.ToString("D2");
            string minutoFinalFormatado = minutofinal.ToString("D2");
            string HoraFinal = $"{horaFinalFormatada}:{minutoFinalFormatado}:00";

            DateTime HoraInicioDT = DateTime.ParseExact(HoraInicio, "HH:mm:ss", null);
            DateTime HoraFinalDT = DateTime.ParseExact(HoraFinal, "HH:mm:ss", null);

            TimeSpan tempoDecorrido = HoraFinalDT - HoraInicioDT;
            TimeSpan tempoAjustado = tempoDecorrido;

            bool horaAlmocoSubtraida = false;

            if (HoraInicioDT.Hour < 12 || (HoraInicioDT.Hour == 12 && HoraInicioDT.Minute < 30))
            {
                if (HoraFinalDT.Hour > 12 || (HoraFinalDT.Hour == 12 && HoraFinalDT.Minute >= 45))
                {
                    TimeSpan subtracao = new TimeSpan(1, 30, 0);
                    tempoAjustado = tempoDecorrido - subtracao;
                    horaAlmocoSubtraida = true;
                }
            }

            string QtdHoras = tempoAjustado.ToString(@"hh\:mm\:ss");
            
            string numeroobra = TextBoxNumeroObra.Text;
            string nomedaobra = labelNomeObra.Text;

            string tarefa = $" Apoio a Obra {nomedaobra} ";
            string Prioridade = "15- Apoio a Obra";
            string Observações = "";

            DateTime DatadaTarefa = DateTimePickerApioaaObra.Value;

            string dataFormatada = DatadaTarefa.Date.ToString("dd/MM/yyyy");


            string nomePreparador = Environment.UserName;
            string[] partes = nomePreparador.Split('.');
            List<string> partesComMaiusculas = new List<string>();

            foreach (string parte in partes)
            {
                if (!string.IsNullOrEmpty(parte))
                {
                    string parteFormatada = char.ToUpper(parte[0]) + parte.Substring(1).ToLower();
                    partesComMaiusculas.Add(parteFormatada);
                }
            }
            string nomeFormatado = string.Join(" ", partesComMaiusculas);

            string CodigodaTarefa = "405";


            ComunicaBD BD = new ComunicaBD();
                 try
                {
                    BD.ConectarBD();
                    string query = "INSERT INTO dbo.RegistoTempo ([Numero da Obra], [Nome da Obra], Tarefa, Preparador, [Hora Inicial], [Hora Final], [Data da Tarefa], [Qtd de Hora], ObservaçõesPreparador, Prioridade, [Codigo da Tarefa]) " +
                                    "VALUES (@NumeroObra, @NomeObra, @Tarefa, @TarefaPreparador, @HoraInicial, @HoraFinal, @DatadaTarefa, @QtddeHora, @ObservaçõesPreparador, @Prioridade, @CodigodaTarefa) ";

                    using (SqlCommand cmd = new SqlCommand(query, BD.GetConnection()))
                    {
                        cmd.Parameters.AddWithValue("@NumeroObra", numeroobra);
                        cmd.Parameters.AddWithValue("@NomeObra", nomedaobra);
                        cmd.Parameters.AddWithValue("@Tarefa", tarefa);
                        cmd.Parameters.AddWithValue("@TarefaPreparador", nomeFormatado);
                        cmd.Parameters.AddWithValue("@HoraInicial", HoraInicio);
                        cmd.Parameters.AddWithValue("@HoraFinal", HoraFinal);
                        cmd.Parameters.AddWithValue("@DatadaTarefa", dataFormatada);
                        cmd.Parameters.AddWithValue("@QtddeHora", QtdHoras);
                        cmd.Parameters.AddWithValue("@ObservaçõesPreparador", Observações);
                        cmd.Parameters.AddWithValue("@Prioridade", Prioridade);
                        cmd.Parameters.AddWithValue("@CodigodaTarefa", CodigodaTarefa);


                        cmd.ExecuteNonQuery();
                    }

                if (horaAlmocoSubtraida)
                {
                    MessageBox.Show($" Hora Registada com sucesso. \n\n Hora Inicio da Tarefa: {HoraInicio} \n\n Hora Terminada da Tarefa: {HoraFinal} \n\n Quantidade de Horas usadas: {QtdHoras} \n\n **A hora de almoço foi descontada.**");
                }
                else
                {
                    MessageBox.Show($" Hora Registada com sucesso. \n\n Hora Inicio da Tarefa: {HoraInicio} \n\n Hora Terminada da Tarefa:: {HoraFinal} \n\n Quantidade de Horas usadas: {QtdHoras}");
                }

            }
                catch (Exception ex)
                {
                    MessageBox.Show("Erro ao registrar a hora de início: " + ex.Message);
                }
                finally
                {
                    BD.DesonectarBD();
                }
            
        }

        public void CarregarNomeObraPorCaminho()
        {
            string anoSelecionado;
            string obraSelecionado;

            obraSelecionado = TextBoxNumeroObra.Text;
            anoSelecionado = "20" + obraSelecionado.Substring(0, 2);
            

            if (!string.IsNullOrEmpty(anoSelecionado) && !string.IsNullOrEmpty(obraSelecionado))
            {
                string caminhoPastaBase = @"\\marconi\COMPANY SHARED FOLDER\OFELIZ\OFM\2.AN\2.CM\DP\1 Obras";

                string restantecaminho = @"1.8 Projeto\1.8.2 Tekla";

                string caminhoPastaAno = Path.Combine(caminhoPastaBase, anoSelecionado, obraSelecionado, restantecaminho);

                if (Directory.Exists(caminhoPastaAno))
                {
                    string[] subpastas = Directory.GetDirectories(caminhoPastaAno);

                    foreach (string subpasta in subpastas)
                    {
                        string nomePasta = Path.GetFileName(subpasta);

                        if (nomePasta.StartsWith(obraSelecionado))
                        {
                            string restanteNome = nomePasta.Substring(obraSelecionado.Length).Trim();

                            string restanteNomeLimpo = RemoverCaracteresEspeciais(restanteNome);

                            labelNomeObra.Text = restanteNomeLimpo;

                            string nomeObraOriginal = labelNomeObra.Text.Replace("_", " ")
                                        .Replace("-", " ")
                                        .Replace("!", " ")
                                        .Replace("@", " ");

                            labelNomeObra.Text = nomeObraOriginal;
                            break;
                        }
                    }
                }
                else
                {
                    MessageBox.Show("O caminho da pasta selecionada não existe.");
                }
            }
            else
            {
                MessageBox.Show("Por favor, selecione um ano e uma obra.");
            }
        }

        private string RemoverCaracteresEspeciais(string input)
        {
            return System.Text.RegularExpressions.Regex.Replace(input, @"[^\w\s]", " ");
        }

        private void ButtonforOpenCalen_Click(object sender, EventArgs e)
        {
            panelCalendario.Visible = !panelCalendario.Visible;
        }

        private void guna2ImageButton1_Click(object sender, EventArgs e)
        {
            panelAlterarDados.Visible = !panelAlterarDados.Visible;
        }

        private void guna2ImageButton2_Click(object sender, EventArgs e)
        {
            SalvarHoraInicialnaBDmanual();
            ConfirmarComunicacaoBD();

        }       

        private void TextBoxAtualizar_TextChanged(object sender, EventArgs e)
        {

        }

        public void ExcluirTarefaSelecionada()
        {
            if (DataGridViewRegistoTempo.SelectedRows.Count > 0)
            {
                int idTarefa = Convert.ToInt32(DataGridViewRegistoTempo.SelectedRows[0].Cells["Id"].Value);

                if (idTarefa != 0)
                {
                    DialogResult result = MessageBox.Show("Tem certeza de que deseja excluir este Registo?", "Confirmar Exclusão", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);

                    if (result == DialogResult.Yes)
                    {
                        string query = "DELETE FROM dbo.RegistoTempo WHERE [Id] = @IdTarefa";

                        ComunicaBD BD = new ComunicaBD();

                        try
                        {
                            BD.ConectarBD();

                            using (SqlCommand cmd = new SqlCommand(query, BD.GetConnection()))
                            {
                                cmd.Parameters.Add("@IdTarefa", SqlDbType.Int).Value = idTarefa;
                                cmd.ExecuteNonQuery();
                            }
                            ConfirmarComunicacaoBD();

                            MessageBox.Show("Registo excluído com sucesso.");
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("Erro ao excluir tarefa: " + ex.Message);
                        }
                        finally
                        {
                            BD.DesonectarBD();
                        }
                    }
                }
                else
                {
                    MessageBox.Show("Tarefa inválida selecionada.");
                }
            }
            else
            {
                MessageBox.Show("Por favor, selecione uma tarefa para excluir.");
            }
        }

        private void guna2ImageButton12_Click(object sender, EventArgs e)
        {
            ExcluirTarefaSelecionada();
        }

        public void CarregarPreparadoresNaComboBox()
        {
            ComunicaBD BD = new ComunicaBD();
            BD.ConectarBD();
            string query = "SELECT Nome FROM dbo.nPreparadores1";
            List<string> list = BD.Procurarbdlist(query);
            BD.DesonectarBD();

            ComboBoxPreparadorAdd.Items.Clear();
            if (list.Count > 0)
            {
                foreach (string nome in list)
                {
                    ComboBoxPreparadorAdd.Items.Add(nome);
                }
            }
            else
            {
                MessageBox.Show("Nenhum preparador encontrado na base de dados.");

            }
        }

        
        public void VerificarUsuario()
        {
            string nomeUsuario = Environment.UserName.ToLower();

            ComunicaBD BD = new ComunicaBD();

            try
            {
                BD.ConectarBD();

                string query = "SELECT AutorizacaoPreparador FROM dbo.nPreparadores1 WHERE [nome.sigla] = @nomeUsuario";

                using (SqlCommand cmd = new SqlCommand(query, BD.GetConnection()))
                {
                    cmd.Parameters.AddWithValue("@nomeUsuario", nomeUsuario);

                    object result = cmd.ExecuteScalar();

                    if (result != null)
                    {
                        bool autorizacao = Convert.ToBoolean(result);

                        if (autorizacao)
                        {
                            label4.Visible = true;
                            ComboBoxPreparadorAdd.Visible = true;
                            ButtonTodos.Visible = true;
                            ButtonTodosComData.Visible = true;

                        }
                        else
                        {
                            label4.Visible = false;
                            ComboBoxPreparadorAdd.Visible = false;
                            ButtonTodos.Visible = false;
                            ButtonTodosComData.Visible = false;

                        }
                    }
                    else
                    {

                        label4.Visible = false;
                        ComboBoxPreparadorAdd.Visible = false;
                        ButtonTodos.Visible = false;
                        ButtonTodosComData.Visible = false;
                    }
                    string nomeUsuario2 = Properties.Settings.Default.NomeUsuario;

                    if (nomeUsuario2 == "ofelizcmadmin" || nomeUsuario2 == "helder.silva")
                    {
                        label4.Visible = true;
                        ComboBoxPreparadorAdd.Visible = true;
                        ButtonTodos.Visible = true;
                        ButtonTodosComData.Visible = true;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erro ao verificar o usuário: " + ex.Message);
            }
            finally
            {
                BD.DesonectarBD();
            }
        }

        private void ButtonTodos_Click(object sender, EventArgs e)
        {
            ComunicaBDparaTabelaTodosSemData();
        }

        private void ComunicaBDparaTabelaTodosSemData()
        {
            ComunicaBD comunicaBD = new ComunicaBD();

            try
            {
                comunicaBD.ConectarBD();

                string Obra = TextBoxNObra.Text;

                string Prioridades = null;

                if (ComboBoxPrioAdd.SelectedItem != null)
                {
                    Prioridades = ComboBoxPrioAdd.SelectedItem.ToString();
                }

                string query = "SELECT Id, [Numero da Obra], [Nome da Obra], Preparador, Tarefa,  [Hora Inicial], [Hora Final], [Data da Tarefa], [Qtd de Hora], Prioridade, ObservaçõesPreparador " +
                               "FROM dbo.RegistoTempo WHERE 1=1";

                if (!string.IsNullOrEmpty(TextBoxNObra.Text))
                {
                    query += " AND [Numero da Obra] = @NumeroObra";
                }

                if (!string.IsNullOrEmpty(Prioridades))
                {
                    query += " AND Prioridade = @Prioridade";
                }

                DataTable dataTable = new DataTable();

                using (var command = new SqlCommand(query, comunicaBD.GetConnection()))
                {
                    if (!string.IsNullOrEmpty(TextBoxNObra.Text))
                    {
                        command.Parameters.AddWithValue("@NumeroObra", Obra);
                    }

                    if (!string.IsNullOrEmpty(Prioridades))
                    {
                        command.Parameters.AddWithValue("@Prioridade", Prioridades);
                    }

                    using (var adapter = new SqlDataAdapter(command))
                    {
                        adapter.Fill(dataTable);
                    }
                }

                DataGridViewRegistoTempo.DataSource = dataTable;
                DataGridViewRegistoTempo.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
                DataGridViewRegistoTempo.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
                DataGridViewRegistoTempo.Columns["Id"].Visible = false;
                DataGridViewRegistoTempo.Columns["Preparador"].Visible = true;
                DataGridViewRegistoTempo.ClearSelection();

                DataGridViewRegistoTempo.Columns["Numero da Obra"].Width = 80;   
                DataGridViewRegistoTempo.Columns["Nome da Obra"].Width = 200;     
                DataGridViewRegistoTempo.Columns["Preparador"].Width = 100;       
                DataGridViewRegistoTempo.Columns["Tarefa"].Width = 200;          
                DataGridViewRegistoTempo.Columns["Hora Inicial"].Width = 80;    
                DataGridViewRegistoTempo.Columns["Hora Final"].Width = 80;      
                DataGridViewRegistoTempo.Columns["Data da Tarefa"].Width = 80;  
                DataGridViewRegistoTempo.Columns["Qtd de Hora"].Width = 80;     
                DataGridViewRegistoTempo.Columns["Prioridade"].Width = 130;      
                DataGridViewRegistoTempo.Columns["ObservaçõesPreparador"].Width = 80;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erro ao conectar à base de dados: " + ex.Message);
            }
            finally
            {
                comunicaBD.DesonectarBD();
            }
        }


        private void ComunicaBDparaTabelaTodosComData()
        {
            ComunicaBD comunicaBD = new ComunicaBD();

            try
            {
                comunicaBD.ConectarBD();

                string Obra = TextBoxNObra.Text;

                string Prioridades = null;

                if (ComboBoxPrioAdd.SelectedItem != null)
                {
                    Prioridades = ComboBoxPrioAdd.SelectedItem.ToString();
                }

                string query = "SELECT Id, [Numero da Obra], [Nome da Obra], Preparador, Tarefa, [Hora Inicial], [Hora Final], [Data da Tarefa], [Qtd de Hora], Prioridade, ObservaçõesPreparador " +
                               "FROM dbo.RegistoTempo WHERE 1=1";

                if (DateTimePickerInicio.Value != DateTimePickerInicio.MinDate)
                {
                    query += " AND TRY_CONVERT(DATETIME, [Data da Tarefa], 103) = @DataInicio";
                }

                if (!string.IsNullOrEmpty(TextBoxNObra.Text))
                {
                    query += " AND [Numero da Obra] = @NumeroObra";
                }

                if (!string.IsNullOrEmpty(Prioridades))
                {
                    query += " AND Prioridade = @Prioridade";
                }                               

                DataTable dataTable = new DataTable();

                using (var command = new SqlCommand(query, comunicaBD.GetConnection()))
                {
                    if (DateTimePickerInicio.Value != DateTimePickerInicio.MinDate)
                    {
                        command.Parameters.AddWithValue("@DataInicio", DateTimePickerInicio.Value.Date);
                    }

                    if (!string.IsNullOrEmpty(TextBoxNObra.Text))
                    {
                        command.Parameters.AddWithValue("@NumeroObra", Obra);
                    }

                    if (!string.IsNullOrEmpty(Prioridades))
                    {
                        command.Parameters.AddWithValue("@Prioridade", Prioridades);
                    }                                    

                    using (var adapter = new SqlDataAdapter(command))
                    {
                        adapter.Fill(dataTable);
                    }
                }

                DataGridViewRegistoTempo.DataSource = dataTable;
                DataGridViewRegistoTempo.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
                DataGridViewRegistoTempo.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
                DataGridViewRegistoTempo.Columns["Id"].Visible = false; 
                DataGridViewRegistoTempo.Columns["Preparador"].Visible = true; 
                DataGridViewRegistoTempo.ClearSelection(); 

                DataGridViewRegistoTempo.Columns["Numero da Obra"].Width = 80;
                DataGridViewRegistoTempo.Columns["Nome da Obra"].Width = 200;
                DataGridViewRegistoTempo.Columns["Preparador"].Width = 100;
                DataGridViewRegistoTempo.Columns["Tarefa"].Width = 200;
                DataGridViewRegistoTempo.Columns["Hora Inicial"].Width = 80;
                DataGridViewRegistoTempo.Columns["Hora Final"].Width = 80;
                DataGridViewRegistoTempo.Columns["Data da Tarefa"].Width = 80;
                DataGridViewRegistoTempo.Columns["Qtd de Hora"].Width = 80;
                DataGridViewRegistoTempo.Columns["Prioridade"].Width = 130;
                DataGridViewRegistoTempo.Columns["ObservaçõesPreparador"].Width = 80;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erro ao conectar à base de dados: " + ex.Message);
            }
            finally
            {
                comunicaBD.DesonectarBD();
            }
        }

        private void ButtonTodosComData_Click(object sender, EventArgs e)
        {
            ComunicaBDparaTabelaTodosComData();
        }

        private void SalvarHoraInicialnaBDmanual()
        {
            if (DataGridViewRegistoTempo.SelectedRows.Count > 0)
            {
                string HoraInicio = DataGridViewRegistoTempo.SelectedRows[0].Cells["Hora Inicial"].Value.ToString();
                string HoraFim = DataGridViewRegistoTempo.SelectedRows[0].Cells["Hora Final"].Value.ToString();
                string NomeTarefa = DataGridViewRegistoTempo.SelectedRows[0].Cells["Tarefa"].Value.ToString();
                string Prioridade = DataGridViewRegistoTempo.SelectedRows[0].Cells["Prioridade"].Value.ToString();
                string ID = DataGridViewRegistoTempo.SelectedRows[0].Cells["ID"].Value.ToString();
                string DataTarefaTerminada = DataGridViewRegistoTempo.SelectedRows[0].Cells["Data da Tarefa"].Value.ToString();

                int horasAlmoco = (int)NumericUpDownHoras.Value;
                int minutosAlmoco = (int)NumericUpDownMinutos.Value;

                TimeSpan subtracao = new TimeSpan(horasAlmoco, minutosAlmoco, 0);

                DateTime horaInicioDT = DateTime.ParseExact(HoraInicio, "HH:mm:ss", null);
                DateTime horaFimDT = DateTime.ParseExact(HoraFim, "HH:mm:ss", null);

                TimeSpan tempoDecorrido = horaFimDT - horaInicioDT;
                TimeSpan tempoAjustado = tempoDecorrido;

                bool horaAlmocoSubtraida = false;

                if (horaInicioDT.Hour < 12 || (horaInicioDT.Hour == 12 && horaInicioDT.Minute < 30))
                {
                    if (horaFimDT.Hour > 12 || (horaFimDT.Hour == 12 && horaFimDT.Minute >= 45))
                    {
                        tempoAjustado = tempoDecorrido - subtracao;
                        horaAlmocoSubtraida = true;
                    }
                }

                string QtdHoras = tempoAjustado.ToString(@"hh\:mm\:ss");

                string query = "UPDATE dbo.RegistoTempo " +
                               "SET [Hora Inicial] = @HoraInicio, [Hora Final] = @HoraFim, [Data da Tarefa] = @DataTarefa, [Qtd de Hora] = @QtdHora " +
                               "WHERE ID = @Id " +
                               "AND [Tarefa] = @Tarefa " +
                               "AND [Prioridade] = @Prioridade";

                ComunicaBD BD = new ComunicaBD();

                try
                {
                    BD.ConectarBD();

                    using (SqlCommand cmd = new SqlCommand(query, BD.GetConnection()))
                    {
                        cmd.Parameters.AddWithValue("@HoraInicio", HoraInicio);
                        cmd.Parameters.AddWithValue("@HoraFim", HoraFim);
                        cmd.Parameters.AddWithValue("@DataTarefa", DataTarefaTerminada);
                        cmd.Parameters.AddWithValue("@QtdHora", QtdHoras);
                        cmd.Parameters.AddWithValue("@Tarefa", NomeTarefa);
                        cmd.Parameters.AddWithValue("@Prioridade", Prioridade);
                        cmd.Parameters.AddWithValue("@Id", ID);

                        cmd.ExecuteNonQuery();
                    }

                    if (horaAlmocoSubtraida)
                    {
                        MessageBox.Show($" Hora Registada com sucesso. \n\n Hora Inicio da Tarefa: {HoraInicio} \n\n Hora Terminada da Tarefa: {HoraFim} \n\n Quantidade de Horas usadas: {QtdHoras} \n\n **A hora de almoço foi descontada.**");
                    }
                    else
                    {
                        MessageBox.Show($" Hora Registada com sucesso. \n\n Hora Inicio da Tarefa: {HoraInicio} \n\n Hora Terminada da Tarefa:: {HoraFim} \n\n Quantidade de Horas usadas: {QtdHoras}");
                    }

                }
                catch (Exception ex)
                {
                    MessageBox.Show("Erro ao registrar as informações: " + ex.Message);
                }
                finally
                {
                    BD.DesonectarBD();
                }
            }
        }


        private TimeSpan ConverterTextoParaTimeSpan(string texto)
        {
            if (TimeSpan.TryParse(texto, out TimeSpan resultado))
            {
                return resultado;
            }

            string[] partes = texto.Split(':');
            int horas = 0, minutos = 0, segundos = 0;

            if (partes.Length == 3) // hh:mm:ss
            {
                int.TryParse(partes[0], out horas);
                int.TryParse(partes[1], out minutos);
                int.TryParse(partes[2], out segundos);
            }
            else if (partes.Length == 2) // mm:ss
            {
                int.TryParse(partes[0], out minutos);
                int.TryParse(partes[1], out segundos);
            }
            else if (partes.Length == 1) // ss
            {
                int.TryParse(partes[0], out segundos);
            }

            return new TimeSpan(horas, minutos, segundos);
        }

        private void DateTimePickerApioaaObra_ValueChanged(object sender, EventArgs e)
        {

        }

        private void guna2ImageButton9_Click(object sender, EventArgs e)
        {
            ConfirmarComunicacaoBD();
            DataGridViewRegistoTempo.ClearSelection();
            TextBoxNObra.Clear();
            ComboBoxPrioAdd.SelectedIndex = -1;
            DateTimePickerInicio.Value = DateTime.Now;
        }

        private void Buttonprimeiro_Click(object sender, EventArgs e)
        {
            ConfirmarComunicacaoBDSemana1();
        }

        private void ButtonSegundo_Click(object sender, EventArgs e)
        {
            ConfirmarComunicacaoBDSemana2();
        }

        private void ButtonTerceiro_Click(object sender, EventArgs e)
        {
            ConfirmarComunicacaoBDSemana3();
        }

        private void ButtonQuarto_Click(object sender, EventArgs e)
        {
            ConfirmarComunicacaoBDSemana4();
        }

        
    }
}
