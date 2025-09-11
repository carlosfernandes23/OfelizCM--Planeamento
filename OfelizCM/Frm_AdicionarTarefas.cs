using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Mail;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using static OfelizCM.PDFCreat;
using MessageBox = System.Windows.Forms.MessageBox;

namespace OfelizCM
{
    public partial class Frm_AdicionarTarefas : Form
    {
        public Frm_AdicionarTarefas()
        {
            InitializeComponent();
            atualizardaComobobox();
            CarregarPastasNaComboBoxAno();
            CarregarPreparadoresNaComboBox();
            CarregarPrioridadesNaComboBox();
            DateTimePickerInicio.CloseUp += DateTimePickerInicio_CloseUp;

        }

        private void Frm_AdicionarTarefas_Load(object sender, EventArgs e)
        {
            Atualizartabaleas();
            DataGridViewAddTarefas.ClearSelection();
            DataGridViewAddTarefas.Focus();
            AtualizarRelatorioTarefas();
            SelcionarDatagriendview();
            DateTimePickerInicio.Value = DateTime.Now;
            DateTimePickerConclusao.Value = DateTime.Now.AddDays(10);
            VerificarUsuario();
        }

        private void Atualizartabaleas()
        {
            ComunicaBDparaTabelaAbertas();
            ComunicaBDparaTabelaPendentes();
            ComunicaBDparaTabelaConcluido();
            OrdenarDataGridViews();

        }

        private void atualizardaComobobox()
        {
            DataGridViewAddTarefas.CellClick += DataGridViewAddTarefas_CellClick;
            DataGridViewPendente.CellClick += DataGridViewPendente_CellClick;
            DataGridViewAguardarAprovação.CellClick += DataGridViewAguardarAprovação_CellClick;

        }

        private void OrdenarDataGridViews()
        {
            if (DataGridViewAddTarefas.Columns.Contains("Data de Conclusão"))
            {
                DataGridViewAddTarefas.Sort(DataGridViewAddTarefas.Columns["Data de Conclusão"], System.ComponentModel.ListSortDirection.Ascending);
            }

            if (DataGridViewPendente.Columns.Contains("Data de Conclusão"))
            {
                DataGridViewPendente.Sort(DataGridViewPendente.Columns["Data de Conclusão"], System.ComponentModel.ListSortDirection.Ascending);
            }

            if (DataGridViewAguardarAprovação.Columns.Contains("Data de Conclusão"))
            {
                DataGridViewAguardarAprovação.Sort(DataGridViewAguardarAprovação.Columns["Data de Conclusão"], System.ComponentModel.ListSortDirection.Ascending);
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
                            ButtonEliminarTarefa.Visible = true;
                            guna2ContainerControl5.Visible = true;
                            guna2HtmlLabel4.Visible = true;
                            guna2ImageButton3.Visible = true;

                        }
                        else
                        {
                            ButtonEliminarTarefa.Visible = false;
                            guna2ContainerControl5.Visible = false;
                            guna2HtmlLabel4.Visible = false;
                            guna2ImageButton3.Visible = false;

                        }
                    }
                    else
                    {

                        ButtonEliminarTarefa.Visible = false;
                        guna2ContainerControl5.Visible = false;
                        guna2HtmlLabel4.Visible = false;
                        guna2ImageButton3.Visible = false;
                    }

                    string nomeUsuario2 = Properties.Settings.Default.NomeUsuario;

                    if (nomeUsuario2 == "ofelizcmadmin" || nomeUsuario2 == "helder.silva")
                    {
                        ButtonEliminarTarefa.Visible = true;
                        guna2ContainerControl5.Visible = true;
                        guna2HtmlLabel4.Visible = true;
                        guna2ImageButton3.Visible = true;
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

        private void ComunicaBDparaTabelaPendentes()
        {
            ComunicaBD comunicaBD = new ComunicaBD();

            try
            {
                comunicaBD.ConectarBD();

                string query = "SELECT Id, [Numero da Obra], [Nome da Obra], Tarefa, Preparador, Estado, Observações, Prioridades, [Data de Inicio], [Data de Conclusão], Concluido , [Data de Conclusão do user] " +
                                       "FROM dbo.RegistoTarefas " +
                                       "WHERE (Estado = 'Pendente') " +
                                       "AND Concluido = 0";

                DataTable dataTable = comunicaBD.Procurarbd(query);

                DataGridViewPendente.DataSource = dataTable;
                DataGridViewPendente.ReadOnly = true;
                DataGridViewPendente.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
                DataGridViewPendente.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
                DataGridViewPendente.Columns["Id"].Visible = false;
                DataGridViewPendente.Columns["Prioridades"].Visible = false;
                DataGridViewPendente.Columns["Concluido"].Visible = false;
                DataGridViewPendente.Columns["Data de Inicio"].Visible = false;
                DataGridViewPendente.Columns["Data de Conclusão"].Visible = false;
                DataGridViewPendente.Columns["Observações"].Width = 70;
                DataGridViewPendente.Columns["Data de Conclusão do user"].Visible = false;
                DataGridViewPendente.Columns["Estado"].Visible = false;
                DataGridViewPendente.ClearSelection();

                DataGridViewPendente.Columns["Numero da Obra"].Width = 50;
                DataGridViewPendente.Columns["Nome da Obra"].Width = 130;
                DataGridViewPendente.Columns["Tarefa"].Width = 200;
                DataGridViewPendente.Columns["Preparador"].Width = 50;
                DataGridViewPendente.Columns["Observações"].Width = 750;

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

        private void ComunicaBDparaTabelaConcluido()
        {
            ComunicaBD comunicaBD = new ComunicaBD();

            try
            {
                comunicaBD.ConectarBD();

                string query = "SELECT Id, [Numero da Obra], [Nome da Obra], Tarefa, Preparador, Estado, Observações, Prioridades, [Data de Inicio], [Data de Conclusão], Concluido , [Data de Conclusão do user] " +
                                       "FROM dbo.RegistoTarefas " +
                                       "WHERE Estado = 'Aguarda aprovação'";

                DataTable dataTable = comunicaBD.Procurarbd(query);

                DataGridViewAguardarAprovação.DataSource = dataTable;
                DataGridViewAguardarAprovação.ReadOnly = true;
                DataGridViewAguardarAprovação.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
                DataGridViewAguardarAprovação.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
                DataGridViewAguardarAprovação.Columns["Id"].Visible = false;
                DataGridViewAguardarAprovação.Columns["Prioridades"].Visible = false;
                DataGridViewAguardarAprovação.Columns["Data de Inicio"].Visible = false;
                DataGridViewAguardarAprovação.Columns["Data de Conclusão"].Visible = false;
                DataGridViewAguardarAprovação.Columns["Concluido"].Visible = false;
                DataGridViewAguardarAprovação.Columns["Data de Conclusão do user"].Visible = false;
                DataGridViewAguardarAprovação.Columns["Estado"].Visible = false;

                DataGridViewAguardarAprovação.ClearSelection();

                DataGridViewAguardarAprovação.Columns["Numero da Obra"].Width = 50;
                DataGridViewAguardarAprovação.Columns["Nome da Obra"].Width = 130;
                DataGridViewAguardarAprovação.Columns["Tarefa"].Width = 200;
                DataGridViewAguardarAprovação.Columns["Preparador"].Width = 50;
                DataGridViewAguardarAprovação.Columns["Observações"].Width = 750;
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

        private void ComunicaBDparaTabelaAbertas()
        {
            ComunicaBD comunicaBD = new ComunicaBD();

            try
            {
                comunicaBD.ConectarBD();

                string query = "SELECT Id, [Numero da Obra], [Nome da Obra], Tarefa, Preparador, Estado, Observações, Prioridades, [Data de Inicio], [Data de Conclusão], Concluido , [Data de Conclusão do user] " +
                                       "FROM dbo.RegistoTarefas " +
                                       "WHERE (Estado IS NULL OR Estado = '') " +
                                       "AND Concluido = 0";

                DataTable dataTable = comunicaBD.Procurarbd(query);

                DataGridViewAddTarefas.DataSource = dataTable;
                DataGridViewAddTarefas.ReadOnly = true;
                DataGridViewAddTarefas.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
                DataGridViewAddTarefas.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
                DataGridViewAddTarefas.Columns["Id"].Visible = false;
                DataGridViewAddTarefas.Columns["Concluido"].Visible = false;
                DataGridViewAddTarefas.Columns["Data de Conclusão do user"].Visible = false;
                DataGridViewAddTarefas.Columns["Observações"].Visible = false;
                DataGridViewAddTarefas.Columns["Estado"].Visible = false;

                DataGridViewAddTarefas.Columns["Numero da Obra"].Width = 25;
                DataGridViewAddTarefas.Columns["Nome da Obra"].Width = 70;
                DataGridViewAddTarefas.Columns["Tarefa"].Width = 120;
                DataGridViewAddTarefas.Columns["Preparador"].Width = 45;
                DataGridViewAddTarefas.Columns["Prioridades"].Width = 60;
                DataGridViewAddTarefas.Columns["Data de Inicio"].Width = 40;
                DataGridViewAddTarefas.Columns["Data de Conclusão"].Width = 100;


                DataGridViewAddTarefas.ClearSelection();
                //DataGridViewAddTarefas.AutoResizeColumns();
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

        public void CarregarPastasNaComboBoxAno()
        {
            string caminhoPasta = @"\\marconi\COMPANY SHARED FOLDER\OFELIZ\OFM\2.AN\2.CM\DP\1 Obras";

            if (Directory.Exists(caminhoPasta))
            {
                string[] subpastas = Directory.GetDirectories(caminhoPasta);

                ComboBoxAnoAdd.Items.Clear();
                string pattern = @"^[A-Za-z0-9]{4}$";
                foreach (string subpasta in subpastas)
                {
                    string nomePasta = Path.GetFileName(subpasta);

                    if (Regex.IsMatch(nomePasta, pattern))
                    {
                        ComboBoxAnoAdd.Items.Add(nomePasta);
                    }
                }
            }
            else
            {
                MessageBox.Show("O caminho especificado não existe.");
            }
        }

        public void CarregarPastasNaComboBoxObras()
        {
            string anoSelecionado = ComboBoxAnoAdd.SelectedItem?.ToString();

            if (!string.IsNullOrEmpty(anoSelecionado))
            {
                string caminhoPastaBase = @"\\marconi\COMPANY SHARED FOLDER\OFELIZ\OFM\2.AN\2.CM\DP\1 Obras\";
                string caminhoPastaAno = Path.Combine(caminhoPastaBase, anoSelecionado);


                if (Directory.Exists(caminhoPastaAno))
                {
                    string[] subpastas = Directory.GetDirectories(caminhoPastaAno);

                    ComboBoxObrasAdd.Items.Clear();

                    foreach (string subpasta in subpastas)
                    {
                        string nomePasta = Path.GetFileName(subpasta);

                        if (nomePasta.Length < 11)
                        {
                            ComboBoxObrasAdd.Items.Add(nomePasta);
                        }
                    }
                }
                else
                {
                    MessageBox.Show("O caminho da pasta selecionada não existe.");
                }
            }
        }

        public void CarregarNomeObraPorCaminho()
        {
            string anoSelecionado;
            string obraSelecionado;

            if (!string.IsNullOrEmpty(TextBoxNObra.Text))
            {
                obraSelecionado = TextBoxNObra.Text;

                anoSelecionado = "20" + obraSelecionado.Substring(0, 2);
            }
            else
            {
                obraSelecionado = ComboBoxObrasAdd.SelectedItem?.ToString();

                anoSelecionado = ComboBoxAnoAdd.SelectedItem?.ToString();
            }

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

        public void InserirTarefaNoBD()
        {
            string nomeObra = labelNomeObra.Text;
            string tarefa = TextBoxTarefaAdd.Text;
            string Preparador = ComboBoxPreparadorAdd.SelectedItem?.ToString();
            string Estado = " ";
            string observacoes = TextBoxObsAdd.Text;
            string prioridades = ComboBoxPrioAdd.SelectedItem?.ToString();
            DateTime dataInicio = DateTimePickerInicio.Value;
            DateTime dataConclusao = DateTimePickerConclusao.Value;
            int Concluido = 0;
            DateTime dataConclusaoUser = guna2DateTimePickerdataconclusaouser.Value;
            int codigodaTarefa = 1;

            string numerodaObra;

            if (!string.IsNullOrEmpty(TextBoxNObra.Text))
            {
                numerodaObra = TextBoxNObra.Text;
            }
            else
            {
                numerodaObra = ComboBoxObrasAdd.SelectedItem?.ToString();
            }


            if (string.IsNullOrEmpty(numerodaObra) || string.IsNullOrEmpty(tarefa) || string.IsNullOrEmpty(Preparador) || string.IsNullOrEmpty(prioridades))
            {
                MessageBox.Show("Por favor, preencha todos os campos obrigatórios.");
                return;
            }

            switch (prioridades)
            {
                case "1- Quantificação Material":
                    codigodaTarefa = 403;
                    break;
                case "2- Modelação Estrutura":
                    codigodaTarefa = 401;
                    break;
                case "3- Modelação Revestimentos":
                    codigodaTarefa = 407;
                    break;
                case "4- Envio para Aprovação 2D/3D Trimble":
                    codigodaTarefa = 402;
                    break;
                case "7- Processo de Fabrico":
                    codigodaTarefa = 403;
                    break;
                case "8- Processo de soldadura":
                    codigodaTarefa = 404;
                    break;
                case "9- Aprovisionamento da parafusaria":
                    codigodaTarefa = 403;
                    break;
                case "10- Enviar IFC para Powerfab":
                    codigodaTarefa = 403;
                    break;
                case "11- Desenhos de Montagem":
                    codigodaTarefa = 413;
                    break;
                case "12- Tarefas Diversas":
                    codigodaTarefa = 408;
                    break;
                case "13- Revisões":
                    codigodaTarefa = 409;
                    break;
                case "14- Apoio a Fabrica":
                    codigodaTarefa = 405;
                    break;
                case "15- Apoio a Obra":
                    codigodaTarefa = 405;
                    break;
                default:
                    codigodaTarefa = 0;
                    break;
            }


            string query = @"
                            INSERT INTO dbo.RegistoTarefas
                            ([Numero da Obra], [Nome da Obra], Tarefa, Preparador, Estado, Observações, Prioridades, [Codigo da Tarefa], [Data de Inicio], [Data de Conclusão], Concluido,  [Data de Conclusão do user])
                            VALUES
                            (@NumerodaObra, @NomedaObra, @TAREFA, @PreparadordaTarefa, @Estado, @Observações, @Prioridades, @CodigodadaTarefa, @DataInicio, @DataConclusão, @Concluido, @DataConclusaoUser)";

            ComunicaBD BD = new ComunicaBD();

            try
            {
                BD.ConectarBD();

                using (SqlCommand cmd = new SqlCommand(query, BD.GetConnection()))
                {
                    cmd.Parameters.AddWithValue("@NumerodaObra", numerodaObra);
                    cmd.Parameters.AddWithValue("@NomedaObra", nomeObra);
                    cmd.Parameters.AddWithValue("@TAREFA", tarefa);
                    cmd.Parameters.AddWithValue("@PreparadordaTarefa", Preparador);
                    cmd.Parameters.AddWithValue("@Estado", Estado);
                    cmd.Parameters.AddWithValue("@Observações", observacoes);
                    cmd.Parameters.AddWithValue("@Prioridades", prioridades);
                    cmd.Parameters.AddWithValue("@CodigodadaTarefa", codigodaTarefa);
                    cmd.Parameters.AddWithValue("@DataInicio", dataInicio);
                    cmd.Parameters.AddWithValue("@DataConclusão", dataConclusao);
                    cmd.Parameters.AddWithValue("@Concluido", Concluido);
                    cmd.Parameters.AddWithValue("@DataConclusaoUser", dataConclusaoUser);

                    cmd.ExecuteNonQuery();
                }

                MessageBox.Show("Tarefa inserida com sucesso!");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erro ao inserir tarefa: " + ex.Message);
            }
            finally
            {
                BD.DesonectarBD();
            }
        }

        public void ExcluirTarefaSelecionada()
        {
            int idTarefa = 0;

            if (DataGridViewAddTarefas.SelectedRows.Count > 0)
            {
                idTarefa = Convert.ToInt32(DataGridViewAddTarefas.SelectedRows[0].Cells["Id"].Value);
            }
            else if (DataGridViewPendente.SelectedRows.Count > 0)
            {
                idTarefa = Convert.ToInt32(DataGridViewPendente.SelectedRows[0].Cells["Id"].Value);
            }
            else if (DataGridViewAguardarAprovação.SelectedRows.Count > 0)
            {
                idTarefa = Convert.ToInt32(DataGridViewAguardarAprovação.SelectedRows[0].Cells["Id"].Value);
            }

            if (idTarefa != 0)
            {
                DialogResult result = MessageBox.Show("Tem certeza de que deseja excluir esta tarefa?", "Confirmar Exclusão", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);

                if (result == DialogResult.Yes)
                {
                    string query = "DELETE FROM dbo.RegistoTarefas WHERE [Id] = @IdTarefa";

                    ComunicaBD BD = new ComunicaBD();

                    try
                    {
                        BD.ConectarBD();

                        using (SqlCommand cmd = new SqlCommand(query, BD.GetConnection()))
                        {
                            cmd.Parameters.AddWithValue("@IdTarefa", idTarefa);

                            cmd.ExecuteNonQuery();
                        }

                        ComunicaBDparaTabelaAbertas();

                        MessageBox.Show("Tarefa excluída com sucesso.");
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
                MessageBox.Show("Por favor, selecione uma tarefa para excluir.");
            }
        }

        private void ComboBoxAnoAdd_SelectedIndexChanged(object sender, EventArgs e)
        {
            CarregarPastasNaComboBoxObras();
        }

        private void ComboBoxObrasAdd_SelectedIndexChanged(object sender, EventArgs e)
        {
            CarregarNomeObraPorCaminho();
        }

        private void ButtonConfirmarTarefa_Click(object sender, EventArgs e)
        {
            CarregarNomeObraPorCaminho();
            InserirTarefaNoBD();
            FiltrarTarefas2();

        }

        private void ButtonIniciarTarefa_Click(object sender, EventArgs e)
        {
            ExcluirTarefaSelecionada();
            FiltrarTarefas2();
        }

        private void ButtonAtualizar_Click(object sender, EventArgs e)
        {
            AtualizarValores();
            FiltrarTarefas2();
        }

        private double CalcularValorRelatorio(DateTime dataConclusao)
        {
            double resultado = 0;

            TimeSpan diferenca = DateTime.Now.Date - dataConclusao.Date;
            resultado = diferenca.Days;

            return resultado;
        }

        private void AtualizarRelatorioTarefas()
        {
            ComunicaBD BD = null;

            try
            {
                BD = new ComunicaBD();
                BD.ConectarBD();

                foreach (DataGridViewRow row in DataGridViewAddTarefas.Rows)
                {
                    if (row.IsNewRow) continue;

                    if (row.Cells["Concluido"].Value != DBNull.Value && (bool)row.Cells["Concluido"].Value == true)
                    {
                        continue;
                    }

                    if (row.Cells["Data de Conclusão"].Value != DBNull.Value)
                    {
                        DateTime? dataConclusao = row.Cells["Data de Conclusão"].Value as DateTime?;

                        if (dataConclusao.HasValue)
                        {
                            double valorRelatorio = CalcularValorRelatorio(dataConclusao.Value);

                            if (valorRelatorio >= 1)
                            {
                                string ID = row.Cells["Id"].Value.ToString();
                                string preparador = row.Cells["Preparador"].Value.ToString();
                                string numeroObra = row.Cells["Numero da Obra"].Value.ToString();
                                string nomeObra = row.Cells["Nome da Obra"].Value.ToString();
                                string tarefa = row.Cells["Tarefa"].Value.ToString();

                                string queryUpdate = "UPDATE dbo.RegistoTarefas " +
                                                     "SET Relatorio = @Relatorio, " +
                                                         "[Data de Conclusão do user] = CAST(GETDATE() AS DATE) " +
                                                     "WHERE Id = @ID AND " +
                                                         "Preparador = @Preparador AND " +
                                                         "[Numero da Obra] = @NumeroObra AND " +
                                                         "[Nome da Obra] = @NomeObra AND " +
                                                         "Tarefa = @Tarefa";

                                using (SqlCommand cmd = new SqlCommand(queryUpdate, BD.GetConnection()))
                                {
                                    cmd.Parameters.AddWithValue("@Relatorio", valorRelatorio);
                                    cmd.Parameters.AddWithValue("@ID", ID);
                                    cmd.Parameters.AddWithValue("@Preparador", preparador);
                                    cmd.Parameters.AddWithValue("@NumeroObra", numeroObra);
                                    cmd.Parameters.AddWithValue("@NomeObra", nomeObra);
                                    cmd.Parameters.AddWithValue("@Tarefa", tarefa);

                                    cmd.ExecuteNonQuery();
                                }
                            }
                            else
                            {
                                continue;
                            }
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("Erro ao atualizar relatórios: " + ex.Message);
            }
            finally
            {
                if (BD != null)
                {
                    BD.DesonectarBD();
                }
            }
        }

        private void SelcionarDatagriendview()
        {
            this.DataGridViewAddTarefas.SelectionChanged += new System.EventHandler(this.DataGridViewAddTarefas_SelectionChanged);
            this.DataGridViewPendente.SelectionChanged += new System.EventHandler(this.DataGridViewPendente_SelectionChanged);
            this.DataGridViewAguardarAprovação.SelectionChanged += new System.EventHandler(this.DataGridViewConcluido_SelectionChanged);
        }

        private void DataGridViewAddTarefas_SelectionChanged(object sender, EventArgs e)
        {
            if (DataGridViewAddTarefas.SelectedRows.Count > 0)
            {
                DataGridViewPendente.ClearSelection();
                DataGridViewAguardarAprovação.ClearSelection();
            }
        }

        private void DataGridViewPendente_SelectionChanged(object sender, EventArgs e)
        {
            if (DataGridViewPendente.SelectedRows.Count > 0)
            {
                DataGridViewAddTarefas.ClearSelection();
                DataGridViewAguardarAprovação.ClearSelection();
            }
        }

        private void DataGridViewConcluido_SelectionChanged(object sender, EventArgs e)
        {
            if (DataGridViewAguardarAprovação.SelectedRows.Count > 0)
            {
                DataGridViewAddTarefas.ClearSelection();
                DataGridViewPendente.ClearSelection();
            }
        }

        private void AtualizarParaPendenteVarios()
        {
            DataGridView dataGridViewSelecionado = null;

            if (DataGridViewAddTarefas.SelectedRows.Count > 0)
            {
                dataGridViewSelecionado = DataGridViewAddTarefas;
            }
            else if (DataGridViewPendente.SelectedRows.Count > 0)
            {
                dataGridViewSelecionado = DataGridViewPendente;
            }
            else if (DataGridViewAguardarAprovação.SelectedRows.Count > 0)
            {
                dataGridViewSelecionado = DataGridViewAguardarAprovação;
            }

            if (dataGridViewSelecionado != null && dataGridViewSelecionado.SelectedRows.Count > 0)
            {
                int idTarefa = Convert.ToInt32(dataGridViewSelecionado.SelectedRows[0].Cells["ID"].Value);
                int concluido = Convert.ToInt32(dataGridViewSelecionado.SelectedRows[0].Cells["Concluido"].Value);

                DialogResult result = MessageBox.Show("Tem certeza de que deseja marcar esta tarefa como 'Pendente'?", "Confirmar Tarefa Pendente", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                if (result == DialogResult.Yes)
                {
                    string query = "UPDATE dbo.RegistoTarefas SET Estado = 'Pendente' ";

                    if (concluido == 1)
                    {
                        query += ", Concluido = 0 ";
                    }
                    query += ",Observações = ' ' ";

                    query += "WHERE ID = @IdTarefa";

                    ComunicaBD BD = new ComunicaBD();

                    try
                    {
                        BD.ConectarBD();

                        using (SqlCommand cmd = new SqlCommand(query, BD.GetConnection()))
                        {
                            cmd.Parameters.AddWithValue("@IdTarefa", idTarefa);
                            cmd.ExecuteNonQuery();

                            Atualizartabaleas();

                            MessageBox.Show("Tarefa marcada como 'Pendente' com sucesso.");
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Erro ao atualizar o estado da tarefa: " + ex.Message);
                    }
                    finally
                    {
                        BD.DesonectarBD();
                    }
                }
            }
            else
            {
                MessageBox.Show("Por favor, selecione uma tarefa para atualizar o estado.");
            }
        }        

        private void AtualizarParaPendenteDefeniçoesOfeliz()
        {
            DataGridView dataGridViewSelecionado = null;

            if (DataGridViewAddTarefas.SelectedRows.Count > 0)
            {
                dataGridViewSelecionado = DataGridViewAddTarefas;
            }
            else if (DataGridViewPendente.SelectedRows.Count > 0)
            {
                dataGridViewSelecionado = DataGridViewPendente;
            }
            else if (DataGridViewAguardarAprovação.SelectedRows.Count > 0)
            {
                dataGridViewSelecionado = DataGridViewAguardarAprovação;
            }

            if (dataGridViewSelecionado != null && dataGridViewSelecionado.SelectedRows.Count > 0)
            {
                int idTarefa = Convert.ToInt32(dataGridViewSelecionado.SelectedRows[0].Cells["ID"].Value);
                int concluido = Convert.ToInt32(dataGridViewSelecionado.SelectedRows[0].Cells["Concluido"].Value);

                DialogResult result = MessageBox.Show("Tem certeza de que deseja marcar esta tarefa como 'Falta Definições P/ OFELIZ'?", "Confirmar Tarefa Pendente", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                if (result == DialogResult.Yes)
                {
                    string query = "UPDATE dbo.RegistoTarefas SET Estado = 'Pendente', Observações = 'Faltam definições da parte do Feliz' ";

                    if (concluido == 1)
                    {
                        query += ", Concluido = 0 ";
                    }

                    query += "WHERE ID = @IdTarefa";

                    ComunicaBD BD = new ComunicaBD();

                    try
                    {
                        BD.ConectarBD();

                        using (SqlCommand cmd = new SqlCommand(query, BD.GetConnection()))
                        {
                            cmd.Parameters.AddWithValue("@IdTarefa", idTarefa);
                            cmd.ExecuteNonQuery();

                            Atualizartabaleas();

                            MessageBox.Show("Tarefa atualizada com sucesso.");
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Erro ao atualizar o estado da tarefa: " + ex.Message);
                    }
                    finally
                    {
                        BD.DesonectarBD();
                    }
                }
            }
            else
            {
                MessageBox.Show("Por favor, selecione uma tarefa para atualizar o estado.");
            }
        }

        private void AtualizarParaPendenteDefeniçoesCliente()
        {
            DataGridView selectedDataGridView = null;
            if (DataGridViewAddTarefas.SelectedRows.Count > 0)
            {
                selectedDataGridView = DataGridViewAddTarefas;
            }
            else if (DataGridViewPendente.SelectedRows.Count > 0)
            {
                selectedDataGridView = DataGridViewPendente;
            }
            else if (DataGridViewAguardarAprovação.SelectedRows.Count > 0)
            {
                selectedDataGridView = DataGridViewAguardarAprovação;
            }

            if (selectedDataGridView != null)
            {
                int idTarefa = Convert.ToInt32(selectedDataGridView.SelectedRows[0].Cells["ID"].Value);
                int concluido = Convert.ToInt32(selectedDataGridView.SelectedRows[0].Cells["Concluido"].Value); // Pega o valor da coluna "Concluido"

                DialogResult result = MessageBox.Show("Tem certeza de que deseja marcar esta tarefa como 'Falta Definições P/ Cliente'?", "Confirmar Tarefa Pendente", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                if (result == DialogResult.Yes)
                {
                    string query = "UPDATE dbo.RegistoTarefas SET Estado = 'Pendente', Observações = 'Faltam definições por parte do cliente' ";

                    if (concluido == 1)
                    {
                        query += ", Concluido = 0 ";
                    }

                    query += "WHERE ID = @IdTarefa";

                    ComunicaBD BD = new ComunicaBD();

                    try
                    {
                        BD.ConectarBD();

                        using (SqlCommand cmd = new SqlCommand(query, BD.GetConnection()))
                        {
                            cmd.Parameters.AddWithValue("@IdTarefa", idTarefa);

                            cmd.ExecuteNonQuery();

                            Atualizartabaleas();

                            MessageBox.Show("Tarefa atualizada com sucesso.");
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Erro ao atualizar o estado da tarefa: " + ex.Message);
                    }
                    finally
                    {
                        BD.DesonectarBD();
                    }
                }
            }
            else
            {
                MessageBox.Show("Por favor, selecione uma tarefa para atualizar o estado.");
            }
        }

        private void guna2Button2_Click_1(object sender, EventArgs e)
        {
            AtualizarParaPendenteVarios();
            Atualizartabaleas();
        }

        private void guna2Button4_Click(object sender, EventArgs e)
        {
            AtualizarParaPendenteDefeniçoesCliente();
            Atualizartabaleas();

        }

        private void guna2Button3_Click(object sender, EventArgs e)
        {
            AtualizarParaPendenteDefeniçoesOfeliz();
            Atualizartabaleas();

        }

        private void guna2Button1_Click(object sender, EventArgs e)
        {

            ConcluirSoldadaduraEmailClara();
            enviarsoldadura();
            AtualizarEstadoTarefa();
            Atualizartabaleas();

        }

        private void AtualizarEstadoTarefa()
        {
            DataGridView dataGridViewSelecionado = null;

            if (DataGridViewAddTarefas.SelectedRows.Count > 0)
            {
                dataGridViewSelecionado = DataGridViewAddTarefas;
            }
            else if (DataGridViewPendente.SelectedRows.Count > 0)
            {
                dataGridViewSelecionado = DataGridViewPendente;
            }
            else if (DataGridViewAguardarAprovação.SelectedRows.Count > 0)
            {
                dataGridViewSelecionado = DataGridViewAguardarAprovação;
            }

            if (dataGridViewSelecionado != null && dataGridViewSelecionado.SelectedRows.Count > 0)
            {
                int idTarefa = Convert.ToInt32(dataGridViewSelecionado.SelectedRows[0].Cells["Id"].Value);
                string Prioridades = dataGridViewSelecionado.SelectedRows[0].Cells["Prioridades"].Value?.ToString();
                int concluido = Convert.ToInt32(dataGridViewSelecionado.SelectedRows[0].Cells["Concluido"].Value);

                DialogResult result = MessageBox.Show("Tem certeza de que deseja marcar esta tarefa como concluída?", "Confirmar Tarefa", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);

                if (result == DialogResult.Yes)
                {
                    string query = "UPDATE dbo.RegistoTarefas SET ";

                    if (concluido == 1)
                    {
                        query += "Concluido = 0, ";
                    }

                    if (Prioridades == "4- Envio para Aprovação 2D/3D Trimble")
                    {
                        query += "Estado = 'Aguarda aprovação', ";
                    }
                    else
                    {
                        query += "Estado = 'Concluído', ";
                    }

                    query += "Observações = ' ', ";

                    query += "[Data de Conclusão do user] = CAST(GETDATE() AS DATE) WHERE Id = @IdTarefa";

                    ComunicaBD BD = new ComunicaBD();

                    try
                    {
                        BD.ConectarBD();

                        using (SqlCommand cmd = new SqlCommand(query, BD.GetConnection()))
                        {
                            cmd.Parameters.AddWithValue("@IdTarefa", idTarefa);
                            cmd.ExecuteNonQuery();
                        }

                        Atualizartabaleas();

                        MessageBox.Show("A tarefa foi marcada como concluída com sucesso.");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Erro ao atualizar o estado da tarefa: " + ex.Message);
                    }
                    finally
                    {
                        BD.DesonectarBD();
                    }
                }
            }
            else
            {
                MessageBox.Show("Por favor, selecione uma tarefa para concluir.");
            }
        }

        private void guna2Button5_Click(object sender, EventArgs e)
        {
            TextBoxNObra.Clear();
            ComboBoxAnoAdd.SelectedIndex = -1;
            ComboBoxObrasAdd.SelectedIndex = -1;
            TextBoxTarefaAdd.Clear();
            TextBoxObsAdd.Clear();
            ComboBoxPreparadorAdd.SelectedIndex = -1;
            ComboBoxPrioAdd.SelectedIndex = -1;
            DateTimePickerInicio.Value = DateTime.Now;
            DateTimePickerConclusao.Value = DateTime.Now.AddDays(10);
        }

        private void FiltrarTarefas()
        {
            string obraSelecionado = null;
            string preparadorSelecionado = null;
            string prioridadeSelecionada = null;

            if (!string.IsNullOrEmpty(TextBoxNObra.Text))
            {
                obraSelecionado = TextBoxNObra.Text;
            }
            else
            {
                obraSelecionado = ComboBoxObrasAdd.SelectedItem?.ToString();
            }

            if (ComboBoxPreparadorAdd.SelectedItem != null)
            {
                preparadorSelecionado = ComboBoxPreparadorAdd.SelectedItem.ToString();
            }

            if (ComboBoxPrioAdd.SelectedItem != null)
            {
                prioridadeSelecionada = ComboBoxPrioAdd.SelectedItem.ToString();
            }

            ComunicaBD BD = new ComunicaBD();
            try
            {
                BD.ConectarBD();

                FiltrarDataGridViewAddTarefas(BD, obraSelecionado, preparadorSelecionado, prioridadeSelecionada);
                FiltrarDataGridViewPendente(BD, obraSelecionado, preparadorSelecionado, prioridadeSelecionada);
                FiltrarDataGridViewAguardarAprovacao(BD, obraSelecionado, preparadorSelecionado, prioridadeSelecionada);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erro ao conectar à base de dados: " + ex.Message);
            }
        }

        private void FiltrarDataGridViewAddTarefas(ComunicaBD BD, string obraSelecionado, string preparadorSelecionado, string prioridadeSelecionada)
        {
            string query = "SELECT * FROM dbo.RegistoTarefas WHERE 1=1";

            if (!string.IsNullOrEmpty(obraSelecionado))
            {
                query += " AND [Numero da Obra] = @NumeroObra";
            }

            if (!string.IsNullOrEmpty(preparadorSelecionado))
            {
                query += " AND Preparador = @Preparador";
            }

            if (!string.IsNullOrEmpty(prioridadeSelecionada))
            {
                query += " AND Prioridades = @Prioridade";
            }

            query += " AND (Estado IS NULL OR Estado = '')";

            using (var command = new SqlCommand(query, BD.GetConnection()))
            {
                if (!string.IsNullOrEmpty(obraSelecionado))
                {
                    command.Parameters.AddWithValue("@NumeroObra", obraSelecionado);
                }

                if (!string.IsNullOrEmpty(preparadorSelecionado))
                {
                    command.Parameters.AddWithValue("@Preparador", preparadorSelecionado);
                }

                if (!string.IsNullOrEmpty(prioridadeSelecionada))
                {
                    command.Parameters.AddWithValue("@Prioridade", prioridadeSelecionada);
                }

                DataTable dataTable = new DataTable();
                using (var adapter = new SqlDataAdapter(command))
                {
                    adapter.Fill(dataTable);
                }

                DataGridViewAddTarefas.DataSource = dataTable;
                DataGridViewAddTarefas.ReadOnly = true;
                DataGridViewAddTarefas.ClearSelection();
                DataGridViewAddTarefas.Columns["Id"].Visible = false;
                DataGridViewAddTarefas.Columns["Concluido"].Visible = false;
                DataGridViewAddTarefas.Columns["Data de Conclusão do user"].Visible = false;
                DataGridViewAddTarefas.Columns["Relatorio"].Visible = false;
                DataGridViewAddTarefas.Columns["Observações do Relatorio"].Visible = false;
                DataGridViewAddTarefas.Columns["Codigo da Tarefa"].Visible = false;
                DataGridViewAddTarefas.Columns["Observações"].Visible = false;
                DataGridViewAddTarefas.Columns["Estado"].Visible = false;
                DataGridViewAddTarefas.AutoResizeColumns();

            }
        }
        
        private void FiltrarDataGridViewPendente(ComunicaBD BD, string obraSelecionado, string preparadorSelecionado, string prioridadeSelecionada)
        {
            string query = "SELECT * FROM dbo.RegistoTarefas WHERE 1=1";

            if (!string.IsNullOrEmpty(obraSelecionado))
            {
                query += " AND [Numero da Obra] = @NumeroObra";
            }

            if (!string.IsNullOrEmpty(preparadorSelecionado))
            {
                query += " AND Preparador = @Preparador";
            }

            if (!string.IsNullOrEmpty(prioridadeSelecionada))
            {
                query += " AND Prioridades = @Prioridade";
            }

            query += " AND Estado = 'Pendente'";

            using (var command = new SqlCommand(query, BD.GetConnection()))
            {
                if (!string.IsNullOrEmpty(obraSelecionado))
                {
                    command.Parameters.AddWithValue("@NumeroObra", obraSelecionado);
                }

                if (!string.IsNullOrEmpty(preparadorSelecionado))
                {
                    command.Parameters.AddWithValue("@Preparador", preparadorSelecionado);
                }

                if (!string.IsNullOrEmpty(prioridadeSelecionada))
                {
                    command.Parameters.AddWithValue("@Prioridade", prioridadeSelecionada);
                }

                DataTable dataTable = new DataTable();
                using (var adapter = new SqlDataAdapter(command))
                {
                    adapter.Fill(dataTable);
                }

                DataGridViewPendente.DataSource = dataTable;
                DataGridViewPendente.ReadOnly = true;
                DataGridViewPendente.ClearSelection();
                DataGridViewPendente.Columns["Id"].Visible = false;
                DataGridViewPendente.Columns["Prioridades"].Visible = false;
                DataGridViewPendente.Columns["Concluido"].Visible = false;
                DataGridViewPendente.Columns["Data de Inicio"].Visible = false;
                DataGridViewPendente.Columns["Data de Conclusão"].Visible = false;
                DataGridViewPendente.Columns["Observações"].Width = 70;
                DataGridViewPendente.Columns["Data de Conclusão do user"].Visible = false;
                DataGridViewPendente.Columns["Relatorio"].Visible = false;
                DataGridViewPendente.Columns["Observações do Relatorio"].Visible = false;
                DataGridViewPendente.Columns["Codigo da Tarefa"].Visible = false;
                DataGridViewPendente.Columns["Estado"].Visible = false;

                DataGridViewPendente.Columns["Numero da Obra"].Width = 50;
                DataGridViewPendente.Columns["Nome da Obra"].Width = 130;
                DataGridViewPendente.Columns["Tarefa"].Width = 200;
                DataGridViewPendente.Columns["Preparador"].Width = 50;
                DataGridViewPendente.Columns["Observações"].Width = 750;
            }
        }

        private void FiltrarDataGridViewAguardarAprovacao(ComunicaBD BD, string obraSelecionado, string preparadorSelecionado, string prioridadeSelecionada)
        {
            string query = "SELECT * FROM dbo.RegistoTarefas WHERE 1=1";

            if (!string.IsNullOrEmpty(obraSelecionado))
            {
                query += " AND [Numero da Obra] = @NumeroObra";
            }

            if (!string.IsNullOrEmpty(preparadorSelecionado))
            {
                query += " AND Preparador = @Preparador";
            }

            if (!string.IsNullOrEmpty(prioridadeSelecionada))
            {
                query += " AND Prioridades = @Prioridade";
            }

            query += " AND Estado = 'Aguarda aprovação'";

            using (var command = new SqlCommand(query, BD.GetConnection()))
            {
                if (!string.IsNullOrEmpty(obraSelecionado))
                {
                    command.Parameters.AddWithValue("@NumeroObra", obraSelecionado);
                }

                if (!string.IsNullOrEmpty(preparadorSelecionado))
                {
                    command.Parameters.AddWithValue("@Preparador", preparadorSelecionado);
                }

                if (!string.IsNullOrEmpty(prioridadeSelecionada))
                {
                    command.Parameters.AddWithValue("@Prioridade", prioridadeSelecionada);
                }

                DataTable dataTable = new DataTable();
                using (var adapter = new SqlDataAdapter(command))
                {
                    adapter.Fill(dataTable);
                }

                DataGridViewAguardarAprovação.DataSource = dataTable;
                DataGridViewAguardarAprovação.ReadOnly = true;
                DataGridViewAguardarAprovação.ClearSelection();
                DataGridViewAguardarAprovação.Columns["Id"].Visible = false;
                DataGridViewAguardarAprovação.Columns["Prioridades"].Visible = false;
                DataGridViewAguardarAprovação.Columns["Data de Inicio"].Visible = false;
                DataGridViewAguardarAprovação.Columns["Data de Conclusão"].Visible = false;
                DataGridViewAguardarAprovação.Columns["Concluido"].Visible = false;
                DataGridViewAguardarAprovação.Columns["Observações"].Width = 70;
                DataGridViewAguardarAprovação.Columns["Data de Conclusão do user"].Visible = false;
                DataGridViewAguardarAprovação.Columns["Relatorio"].Visible = false;
                DataGridViewAguardarAprovação.Columns["Observações do Relatorio"].Visible = false;
                DataGridViewAguardarAprovação.Columns["Codigo da Tarefa"].Visible = false;
                DataGridViewAguardarAprovação.Columns["Estado"].Visible = false;

                DataGridViewAguardarAprovação.Columns["Numero da Obra"].Width = 50;
                DataGridViewAguardarAprovação.Columns["Nome da Obra"].Width = 130;
                DataGridViewAguardarAprovação.Columns["Tarefa"].Width = 200;
                DataGridViewAguardarAprovação.Columns["Preparador"].Width = 50;
                DataGridViewAguardarAprovação.Columns["Observações"].Width = 750;
            }
        }    

        private void AtualizarValores()
        {
            if (DataGridViewAddTarefas.SelectedRows.Count > 0)
            {
                AtualizarLinha(DataGridViewAddTarefas);
            }
            else if (DataGridViewPendente.SelectedRows.Count > 0)
            {
                AtualizarLinha(DataGridViewPendente);
            }
            else if (DataGridViewAguardarAprovação.SelectedRows.Count > 0)
            {
                AtualizarLinha(DataGridViewAguardarAprovação);
            }
            else
            {
                MessageBox.Show("Selecione uma linha para atualizar.");
            }
        }

        private void AtualizarLinha(DataGridView dataGridView)
        {
            var selectedRow = dataGridView.SelectedRows[0];

            if (selectedRow == null)
            {
                MessageBox.Show("Nenhuma linha selecionada.");
                return;
            }

            int id = Convert.ToInt32(selectedRow.Cells["Id"].Value);

            string obraSelecionado = null;
            if (!string.IsNullOrEmpty(TextBoxNObra.Text))
            {
                obraSelecionado = TextBoxNObra.Text;
            }
            else if (ComboBoxObrasAdd.SelectedItem != null)
            {
                obraSelecionado = ComboBoxObrasAdd.SelectedItem.ToString();
            }

            string preparadorSelecionado = ComboBoxPreparadorAdd.SelectedItem?.ToString();
            string prioridadeSelecionada = ComboBoxPrioAdd.SelectedItem?.ToString();
            string estadoSelecionado = ComboBoxEstado.SelectedItem?.ToString();
            string tarefa = TextBoxTarefaAdd.Text;
            string nomeObra = labelNomeObra.Text;
            string observacoes = TextBoxObsAdd.Text;
            DateTime? dataInicio = DateTimePickerInicio.Value;
            DateTime? dataConclusao = DateTimePickerConclusao.Value;

            DateTime dataConclusaoLinha;
            bool dataConclusaoAlterada = false;

            var dataConclusaoCelula = selectedRow.Cells["Data de Conclusão"].Value;

            if (dataConclusaoCelula != null && DateTime.TryParse(dataConclusaoCelula.ToString(), out dataConclusaoLinha))
            {
                if (dataConclusao.HasValue && dataConclusao.Value != dataConclusaoLinha)
                {
                    dataConclusaoAlterada = true;
                }
            }

            ComunicaBD BD = new ComunicaBD();
            try
            {
                BD.ConectarBD();

                string query = "UPDATE dbo.RegistoTarefas SET ";
                List<SqlParameter> parametros = new List<SqlParameter>();

                if (!string.IsNullOrEmpty(nomeObra))
                {
                    query += "[Nome da Obra] = @NomeObra, ";
                    parametros.Add(new SqlParameter("@NomeObra", nomeObra));
                }

                if (!string.IsNullOrEmpty(obraSelecionado))
                {
                    query += "[Numero da Obra] = @NumerodaObra, ";
                    parametros.Add(new SqlParameter("@NumerodaObra", obraSelecionado));
                }

                if (!string.IsNullOrEmpty(preparadorSelecionado))
                {
                    query += "Preparador = @Preparador, ";
                    parametros.Add(new SqlParameter("@Preparador", preparadorSelecionado));
                }

                if (!string.IsNullOrEmpty(prioridadeSelecionada))
                {
                    query += "Prioridades = @Prioridades, ";
                    parametros.Add(new SqlParameter("@Prioridades", prioridadeSelecionada));
                }

                if (!string.IsNullOrEmpty(tarefa))
                {
                    query += "Tarefa = @Tarefa, ";
                    parametros.Add(new SqlParameter("@Tarefa", tarefa));
                }

                if (estadoSelecionado == "Em Execução" && !string.IsNullOrEmpty(observacoes))
                {
                    query += "Observações = @Observacoes, ";
                    parametros.Add(new SqlParameter("@Observacoes", observacoes));
                }
                else if (estadoSelecionado != "Em Execução" && !string.IsNullOrEmpty(observacoes))
                {
                    query += "Observações = @Observacoes, ";
                    parametros.Add(new SqlParameter("@Observacoes", observacoes));
                }
                else if (estadoSelecionado == "Em Execução")
                {
                    query += "Observações = '', ";
                }
                else
                {
                    query += "Observações = NULL, ";
                }

                if (dataInicio.HasValue)
                {
                    query += "[Data de Inicio] = @DataInicio, ";
                    parametros.Add(new SqlParameter("@DataInicio", dataInicio.Value));
                }

                if (dataConclusao.HasValue)
                {
                    query += "[Data de Conclusão] = @DataConclusao, ";
                    parametros.Add(new SqlParameter("@DataConclusao", dataConclusao.Value));
                }

                if (estadoSelecionado == "Em Execução")
                {
                    query += "Estado = @Estado, Concluido = 0, ";

                    parametros.Add(new SqlParameter("@Estado", ""));
                }
                else if (!string.IsNullOrEmpty(estadoSelecionado))
                {
                    query += "Estado = @Estado, Concluido = 0, ";
                    parametros.Add(new SqlParameter("@Estado", estadoSelecionado));
                }

                if (dataConclusaoAlterada)
                {
                    query += "Relatorio = 0, ";
                }

                if (query.EndsWith(", "))
                {
                    query = query.Substring(0, query.Length - 2);
                }

                query += " WHERE Id = @Id";
                parametros.Add(new SqlParameter("@Id", id));

                using (var command = new SqlCommand(query, BD.GetConnection()))
                {
                    command.Parameters.AddRange(parametros.ToArray());

                    int linhasAfetadas = command.ExecuteNonQuery();
                    if (linhasAfetadas > 0)
                    {
                        MessageBox.Show("Dados atualizados com sucesso.");
                    }
                    else
                    {
                        MessageBox.Show("Nenhuma linha foi atualizada.");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erro ao atualizar os dados: " + ex.Message);
            }
        }

        private void AtualizarValoresConcluido()
        {
            if (DataGridViewAddTarefas.SelectedRows.Count > 0)
            {
                AtualizarLinhaConcluido(DataGridViewAddTarefas);
            }
            else if (DataGridViewPendente.SelectedRows.Count > 0)
            {
                AtualizarLinhaConcluido(DataGridViewPendente);
            }
            else if (DataGridViewAguardarAprovação.SelectedRows.Count > 0)
            {
                AtualizarLinhaConcluido(DataGridViewAguardarAprovação);
            }
            else
            {
                MessageBox.Show("Selecione uma linha para atualizar.");
            }
        }

        private void AtualizarLinhaConcluido(DataGridView dataGridView)
        {
            if (dataGridView.SelectedRows.Count == 0)
            {
                MessageBox.Show("Por favor, selecione uma tarefa para concluir.");
                return;
            }
            var selectedRow = dataGridView.SelectedRows[0];
            if (selectedRow == null)
            {
                MessageBox.Show("Nenhuma linha selecionada.");
                return;
            }

            int id = Convert.ToInt32(selectedRow.Cells["Id"].Value);

            string prioridades = selectedRow.Cells["Prioridades"].Value.ToString();

            DialogResult result = MessageBox.Show("Tem certeza de que deseja marcar esta tarefa como concluída?", "Confirmar Tarefa", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);

            if (result == DialogResult.Yes)
            {
                string query;

                if (prioridades == "4- Envio para Aprovação 2D/3D Trimble")
                {
                    query = "UPDATE dbo.RegistoTarefas SET Concluido = 1, Estado = 'Aguarda aprovação', [Data de Conclusão do user] = CAST(GETDATE() AS DATE) WHERE Id = @IdTarefa";
                }
                else
                {
                    query = "UPDATE dbo.RegistoTarefas SET Concluido = 1, Estado = 'Concluído', [Data de Conclusão do user] = CAST(GETDATE() AS DATE) WHERE Id = @IdTarefa";
                }

                ComunicaBD BD = new ComunicaBD();
                try
                {
                    BD.ConectarBD();

                    using (SqlCommand cmd = new SqlCommand(query, BD.GetConnection()))
                    {
                        cmd.Parameters.AddWithValue("@IdTarefa", id);
                        cmd.ExecuteNonQuery();
                    }

                    MessageBox.Show("A tarefa foi marcada como concluída com sucesso.");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Erro ao atualizar o estado da tarefa: " + ex.Message);
                }
                finally
                {
                    BD.DesonectarBD();
                }
            }
        }

        private void guna2ImageButton3_Click(object sender, EventArgs e)
        {
            string folderPath = @"C:\r";

            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }

            string filePath = Path.Combine(folderPath, "relatorio.pdf");

            ExportDataGridViewToPdf export = new ExportDataGridViewToPdf(DataGridViewAddTarefas);
            export.ExportToPdf(filePath);

            try
            {
                if (File.Exists(filePath))
                {
                    Process.Start(filePath);
                }
                else
                {
                    MessageBox.Show("O arquivo PDF não foi gerado corretamente.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erro ao tentar abrir o arquivo PDF: " + ex.Message);
            }
        }

        private void DataGridViewAddTarefas_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                DataGridViewRow row = DataGridViewAddTarefas.Rows[e.RowIndex];

                TextBoxNObra.Text = row.Cells["Numero da Obra"].Value.ToString();
                labelNomeObra.Text = row.Cells["Nome da Obra"].Value.ToString();
                TextBoxTarefaAdd.Text = row.Cells["Tarefa"].Value.ToString();
                ComboBoxPreparadorAdd.SelectedItem = row.Cells["Preparador"].Value.ToString();
                ComboBoxPrioAdd.SelectedItem = row.Cells["Prioridades"].Value.ToString();
                ComboBoxEstado.SelectedItem = "Em Execução";
                TextBoxObsAdd.Text = row.Cells["Observações"].Value.ToString();
                DateTimePickerInicio.Value = Convert.ToDateTime(row.Cells["Data de Inicio"].Value);
                DateTimePickerConclusao.Value = Convert.ToDateTime(row.Cells["Data de Conclusão"].Value);
            }
        }

        private void DataGridViewPendente_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                DataGridViewRow row = DataGridViewPendente.Rows[e.RowIndex];

                TextBoxNObra.Text = row.Cells["Numero da Obra"].Value.ToString();
                labelNomeObra.Text = row.Cells["Nome da Obra"].Value.ToString();
                TextBoxTarefaAdd.Text = row.Cells["Tarefa"].Value.ToString();
                ComboBoxPreparadorAdd.SelectedItem = row.Cells["Preparador"].Value.ToString();
                ComboBoxPrioAdd.SelectedItem = row.Cells["Prioridades"].Value.ToString();
                ComboBoxEstado.SelectedItem = "Pendente";
                TextBoxObsAdd.Text = row.Cells["Observações"].Value.ToString();
                DateTimePickerInicio.Value = Convert.ToDateTime(row.Cells["Data de Inicio"].Value);
                DateTimePickerConclusao.Value = Convert.ToDateTime(row.Cells["Data de Conclusão"].Value);
            }
        }

        private void DataGridViewAguardarAprovação_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                DataGridViewRow row = DataGridViewAguardarAprovação.Rows[e.RowIndex];

                TextBoxNObra.Text = row.Cells["Numero da Obra"].Value.ToString();
                labelNomeObra.Text = row.Cells["Nome da Obra"].Value.ToString();
                TextBoxTarefaAdd.Text = row.Cells["Tarefa"].Value.ToString();
                ComboBoxPreparadorAdd.SelectedItem = row.Cells["Preparador"].Value.ToString();
                ComboBoxPrioAdd.SelectedItem = row.Cells["Prioridades"].Value.ToString();
                ComboBoxEstado.SelectedItem = "Aguarda aprovação";
                TextBoxObsAdd.Text = row.Cells["Observações"].Value.ToString();
                DateTimePickerInicio.Value = Convert.ToDateTime(row.Cells["Data de Inicio"].Value);
                DateTimePickerConclusao.Value = Convert.ToDateTime(row.Cells["Data de Conclusão"].Value);
            }
        }

        private void guna2ImageButton13_Click(object sender, EventArgs e)
        {
            AtualizarParaPendenteVarios();
            Atualizartabaleas();
        }

        private void guna2ImageButton8_Click(object sender, EventArgs e)
        {
            AtualizarParaPendenteDefeniçoesOfeliz();
            Atualizartabaleas();
        }

        private void guna2ImageButton7_Click(object sender, EventArgs e)
        {
            AtualizarParaPendenteDefeniçoesCliente();
            Atualizartabaleas();
        }

        private void guna2ImageButton2_Click(object sender, EventArgs e)
        {
            AtualizarValoresConcluido();
            Atualizartabaleas();
        }
                
        private void FiltrarTarefas2()
        {
            string PreparadorSelecionado = null;



            PreparadorSelecionado = ComboBoxPreparadorAdd.SelectedItem?.ToString();


            ComunicaBD BD = new ComunicaBD();
            try
            {
                BD.ConectarBD();

                FiltrarDataGridViewAddTarefas2(BD, PreparadorSelecionado);
                FiltrarDataGridViewPendente2(BD, PreparadorSelecionado);
                FiltrarDataGridViewAguardarAprovacao2(BD, PreparadorSelecionado);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erro ao conectar à base de dados: " + ex.Message);
            }
        }

        private void FiltrarDataGridViewAddTarefas2(ComunicaBD BD, string PreparadorSelecionado)
        {
            string query = "SELECT * FROM dbo.RegistoTarefas WHERE 1=1";

            if (!string.IsNullOrEmpty(PreparadorSelecionado))
            {
                query += " AND Preparador = @Preparador";
            }


            query += " AND (Estado IS NULL OR Estado = '')";

            using (var command = new SqlCommand(query, BD.GetConnection()))
            {
                if (!string.IsNullOrEmpty(PreparadorSelecionado))
                {
                    command.Parameters.AddWithValue("@Preparador", PreparadorSelecionado);
                }


                DataTable dataTable = new DataTable();
                using (var adapter = new SqlDataAdapter(command))
                {
                    adapter.Fill(dataTable);
                }

                DataGridViewAddTarefas.DataSource = dataTable;
                DataGridViewAddTarefas.ReadOnly = true;
                DataGridViewAddTarefas.ClearSelection();
                DataGridViewAddTarefas.Columns["Id"].Visible = false;
                DataGridViewAddTarefas.Columns["Concluido"].Visible = false;
                DataGridViewAddTarefas.Columns["Data de Conclusão do user"].Visible = false;
                DataGridViewAddTarefas.Columns["Relatorio"].Visible = false;
                DataGridViewAddTarefas.Columns["Observações do Relatorio"].Visible = false;
                DataGridViewAddTarefas.Columns["Codigo da Tarefa"].Visible = false;
                DataGridViewAddTarefas.Columns["Observações"].Visible = false;
                DataGridViewAddTarefas.Columns["Estado"].Visible = false;
                DataGridViewAddTarefas.AutoResizeColumns();

            }
        }

        private void FiltrarDataGridViewPendente2(ComunicaBD BD, string PreparadorSelecionado)
        {
            string query = "SELECT * FROM dbo.RegistoTarefas WHERE 1=1";

            if (!string.IsNullOrEmpty(PreparadorSelecionado))
            {
                query += " AND Preparador = @Preparador";
            }

            query += " AND Estado = 'Pendente'";

            using (var command = new SqlCommand(query, BD.GetConnection()))
            {
                if (!string.IsNullOrEmpty(PreparadorSelecionado))
                {
                    command.Parameters.AddWithValue("@Preparador", PreparadorSelecionado);
                }

                DataTable dataTable = new DataTable();
                using (var adapter = new SqlDataAdapter(command))
                {
                    adapter.Fill(dataTable);
                }

                DataGridViewPendente.DataSource = dataTable;
                DataGridViewPendente.ReadOnly = true;
                DataGridViewPendente.ClearSelection();
                DataGridViewPendente.Columns["Id"].Visible = false;
                DataGridViewPendente.Columns["Prioridades"].Visible = false;
                DataGridViewPendente.Columns["Concluido"].Visible = false;
                DataGridViewPendente.Columns["Data de Inicio"].Visible = false;
                DataGridViewPendente.Columns["Data de Conclusão"].Visible = false;
                DataGridViewPendente.Columns["Data de Conclusão do user"].Visible = false;
                DataGridViewPendente.Columns["Relatorio"].Visible = false;
                DataGridViewPendente.Columns["Observações do Relatorio"].Visible = false;
                DataGridViewPendente.Columns["Codigo da Tarefa"].Visible = false;
                DataGridViewPendente.Columns["Estado"].Visible = false;

                DataGridViewPendente.Columns["Numero da Obra"].Width = 50;
                DataGridViewPendente.Columns["Nome da Obra"].Width = 130;
                DataGridViewPendente.Columns["Tarefa"].Width = 200;
                DataGridViewPendente.Columns["Preparador"].Width = 50;
                DataGridViewPendente.Columns["Observações"].Width = 750;

            }
        }

        private void FiltrarDataGridViewAguardarAprovacao2(ComunicaBD BD, string PreparadorSelecionado)
        {
            string query = "SELECT * FROM dbo.RegistoTarefas WHERE 1=1";

            if (!string.IsNullOrEmpty(PreparadorSelecionado))
            {
                query += " AND Preparador = @Preparador";
            }

            query += " AND Estado = 'Aguarda aprovação'";

            using (var command = new SqlCommand(query, BD.GetConnection()))
            {
                if (!string.IsNullOrEmpty(PreparadorSelecionado))
                {
                    command.Parameters.AddWithValue("@Preparador", PreparadorSelecionado);
                }

                DataTable dataTable = new DataTable();
                using (var adapter = new SqlDataAdapter(command))
                {
                    adapter.Fill(dataTable);
                }

                DataGridViewAguardarAprovação.DataSource = dataTable;
                DataGridViewAguardarAprovação.ReadOnly = true;
                DataGridViewAguardarAprovação.ClearSelection();
                DataGridViewAguardarAprovação.Columns["Id"].Visible = false;
                DataGridViewAguardarAprovação.Columns["Prioridades"].Visible = false;
                DataGridViewAguardarAprovação.Columns["Data de Inicio"].Visible = false;
                DataGridViewAguardarAprovação.Columns["Data de Conclusão"].Visible = false;
                DataGridViewAguardarAprovação.Columns["Concluido"].Visible = false;
                DataGridViewAguardarAprovação.Columns["Observações"].Width = 70;
                DataGridViewAguardarAprovação.Columns["Data de Conclusão do user"].Visible = false;
                DataGridViewAguardarAprovação.Columns["Relatorio"].Visible = false;
                DataGridViewAguardarAprovação.Columns["Observações do Relatorio"].Visible = false;
                DataGridViewAguardarAprovação.Columns["Codigo da Tarefa"].Visible = false;
                DataGridViewAguardarAprovação.Columns["Estado"].Visible = false;

                DataGridViewAguardarAprovação.Columns["Numero da Obra"].Width = 50;
                DataGridViewAguardarAprovação.Columns["Nome da Obra"].Width = 130;
                DataGridViewAguardarAprovação.Columns["Tarefa"].Width = 200;
                DataGridViewAguardarAprovação.Columns["Preparador"].Width = 50;
                DataGridViewAguardarAprovação.Columns["Observações"].Width = 750;
            }
        }            

        private void guna2ImageButton1_Click(object sender, EventArgs e)
        {
            FiltrarTarefas();
        }

        private void guna2ImageButton9_Click(object sender, EventArgs e)
        {
            TextBoxNObra.Clear();
            ComboBoxAnoAdd.SelectedIndex = -1;
            ComboBoxObrasAdd.SelectedIndex = -1;
            TextBoxTarefaAdd.Clear();
            TextBoxObsAdd.Clear();
            ComboBoxEstado.SelectedIndex = -1;
            ComboBoxPreparadorAdd.SelectedIndex = -1;
            ComboBoxPrioAdd.SelectedIndex = -1;
            DateTimePickerInicio.Value = DateTime.Now;
            DateTimePickerConclusao.Value = DateTime.Now.AddDays(10);
            Atualizartabaleas();
        }

        private void guna2ImageButton12_Click(object sender, EventArgs e)
        {
            guna2ImageButton3.PerformClick();
            guna2ImageButton14.PerformClick();
            guna2ImageButton15.PerformClick();
        }

        private void guna2ImageButton4_Click(object sender, EventArgs e)
        {
            guna2ImageButton16.PerformClick();
            CarregarNomeObraPorCaminho();
            VerificarTarefaExistente();
            InserirTarefaNoBD();
            FiltrarTarefas2();
        }

        private void guna2ImageButton5_Click(object sender, EventArgs e)
        {
            guna2ImageButton16.PerformClick();
            CarregarNomeObraPorCaminho();
            AtualizarValores();
            FiltrarTarefas2();
        }

        private void guna2ImageButton11_Click(object sender, EventArgs e)
        {
            TextBoxNObra.Clear();
            ComboBoxAnoAdd.SelectedIndex = -1;
            ComboBoxObrasAdd.SelectedIndex = -1;
            TextBoxTarefaAdd.Clear();
            TextBoxObsAdd.Clear();
            ComboBoxPreparadorAdd.SelectedIndex = -1;
            ComboBoxPrioAdd.SelectedIndex = -1;
            DateTimePickerInicio.Value = DateTime.Now;
            DateTimePickerConclusao.Value = DateTime.Now.AddDays(10);
        }

        private void guna2ImageButton6_Click(object sender, EventArgs e)
        {
            if (ExisteTarefaEmAndamento())
            {
                return;
            }
            ExcluirTarefaSelecionada();
            FiltrarTarefas2();
        }

        private void guna2ImageButton10_Click(object sender, EventArgs e)
        {
            FiltrarTarefas();
        }

        private void DateTimePickerInicio_CloseUp(object sender, EventArgs e)
        {
            DateTimePickerConclusao.Value = DateTimePickerInicio.Value;
        }

        private void guna2ImageButton14_Click(object sender, EventArgs e)
        {
            string folderPath = @"C:\r";

            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }

            string filePath = Path.Combine(folderPath, "relatoriopendente.pdf");

            ExportDataGridViewToPdf export = new ExportDataGridViewToPdf(DataGridViewPendente);
            export.ExportToPdf(filePath);

            try
            {
                if (File.Exists(filePath))
                {
                    Process.Start(filePath);
                }
                else
                {
                    MessageBox.Show("O arquivo PDF não foi gerado corretamente.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erro ao tentar abrir o arquivo PDF: " + ex.Message);
            }
        }

        private void guna2ImageButton15_Click(object sender, EventArgs e)
        {
            string folderPath = @"C:\r";

            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }

            string filePath = Path.Combine(folderPath, "relatorioaguardaraprovacao.pdf");

            ExportDataGridViewToPdf export = new ExportDataGridViewToPdf(DataGridViewAguardarAprovação);
            export.ExportToPdf(filePath);

            try
            {
                if (File.Exists(filePath))
                {
                    Process.Start(filePath);
                }
                else
                {
                    MessageBox.Show("O arquivo PDF não foi gerado corretamente.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erro ao tentar abrir o arquivo PDF: " + ex.Message);
            }
        }

        private async void guna2ImageButton16_Click(object sender, EventArgs e)
        {
            string texto = TextBoxObsAdd.Text;

            if (!string.IsNullOrEmpty(texto))
            {
                string textoCorrigido = await VerificarOrtografiaComLanguageTool(texto);

                if (!string.IsNullOrEmpty(textoCorrigido))
                {
                    TextBoxObsAdd.Text = textoCorrigido;
                }
            }
            else
            {
            }
        }

        static async Task<string> VerificarOrtografiaComLanguageTool(string texto)
        {
            string url = "https://api.languagetool.org/v2/check";
            var client = new HttpClient();

            var content = new StringContent($"text={texto}&language=pt-BR", Encoding.UTF8, "application/x-www-form-urlencoded");

            try
            {
                var response = await client.PostAsync(url, content);

                if (response.IsSuccessStatusCode)
                {
                    var respostaJson = await response.Content.ReadAsStringAsync();

                    var json = JObject.Parse(respostaJson);

                    foreach (var erro in json["matches"])
                    {
                        string erroTexto = erro["message"].ToString();
                        int inicioErro = (int)erro["offset"];
                        int fimErro = inicioErro + (int)erro["length"];

                        string sugestao = erro["replacements"]?.FirstOrDefault()?["value"]?.ToString();

                        if (!string.IsNullOrEmpty(sugestao))
                        {
                            texto = texto.Substring(0, inicioErro) + sugestao + texto.Substring(fimErro);
                        }
                    }

                    return texto;
                }
                else
                {
                    MessageBox.Show("Erro ao verificar a ortografia. Tente novamente.");
                    return texto;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao acessar a API: {ex.Message}");
                return texto;
            }
        }

        private void enviarsoldadura()
        {
            if (DataGridViewAddTarefas.SelectedRows.Count > 0)
            {
                int idTarefa = Convert.ToInt32(DataGridViewAddTarefas.SelectedRows[0].Cells["Id"].Value);
                string Prioridades = DataGridViewAddTarefas.SelectedRows[0].Cells["Prioridades"].Value?.ToString();
                string Preparador = DataGridViewAddTarefas.SelectedRows[0].Cells["Preparador"].Value?.ToString();

                if (Preparador == "Helder Silva" || Preparador == "Elias Tinoco" || Prioridades == "8- Processo de soldadura" || Preparador == "Carlos Alves")
                {
                    AtualizarEstadoTarefaSoldadura("Elias Tinoco");
                    AtualizarEstadoTarefaSoldadura("Helder Silva");

                }
                else { }
            }
        }

        private void AtualizarEstadoTarefaSoldadura(string preparador)
        {
            if (DataGridViewAddTarefas.SelectedRows.Count > 0)
            {
                int idTarefa = Convert.ToInt32(DataGridViewAddTarefas.SelectedRows[0].Cells["Id"].Value);
                string Prioridades = DataGridViewAddTarefas.SelectedRows[0].Cells["Prioridades"].Value?.ToString();
                string Numeroobra = DataGridViewAddTarefas.SelectedRows[0].Cells["Numero da Obra"].Value?.ToString();
                string Tarefa = DataGridViewAddTarefas.SelectedRows[0].Cells["Tarefa"].Value?.ToString();
                string userName = Environment.UserName;

                if (userName == "helder.silva" || userName == "elias.tinoco" || Prioridades == "8- Processo de soldadura" || userName == "carlos.alves")
                {
                    string query = "UPDATE dbo.RegistoTarefas SET Concluido = 1, Estado = 'Concluído', [Data de Conclusão do user] = CAST(GETDATE() AS DATE) " +
                                   "WHERE Prioridades = @Prioridades AND [Numero da Obra] = @NumerodaObra AND Tarefa = @Tarefa AND Preparador = @Preparador";

                    ComunicaBD BD = new ComunicaBD();

                    try
                    {
                        BD.ConectarBD();

                        using (SqlCommand cmd = new SqlCommand(query, BD.GetConnection()))
                        {
                            cmd.Parameters.AddWithValue("@Prioridades", Prioridades);
                            cmd.Parameters.AddWithValue("@NumerodaObra", Numeroobra);
                            cmd.Parameters.AddWithValue("@Tarefa", Tarefa);
                            cmd.Parameters.AddWithValue("@Preparador", preparador);

                            cmd.ExecuteNonQuery();
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Erro ao atualizar o estado da tarefa: " + ex.Message);
                    }
                    finally
                    {
                        BD.DesonectarBD();
                    }
                }
                else
                {
                    MessageBox.Show("Você não tem permissão para atualizar esta tarefa.");
                }
            }
            else
            {
                MessageBox.Show("Nenhuma tarefa selecionada.");
            }
        }

        private string GetSaudacao()
        {
            DateTime horaAtual = DateTime.Now;
            if (horaAtual.Hour < 12 || (horaAtual.Hour == 12 && horaAtual.Minute < 30))
            {
                return "Bom Dia,";
            }
            else
            {
                return "Boa Tarde,";
            }
        }

        private void ConcluirSoldadaduraEmailClara()
        {
            string Prioridades = DataGridViewAddTarefas.SelectedRows[0].Cells["Tarefa"].Value?.ToString();
            string Nomeobra = DataGridViewAddTarefas.SelectedRows[0].Cells["Nome da Obra"].Value?.ToString();
            string NumeroObra = DataGridViewAddTarefas.SelectedRows[0].Cells["Numero da Obra"].Value?.ToString();

            if (Prioridades.Contains("Processo de Soldadura"))
            {
                string PrioridadeSomenteNumeros = Regex.Replace(Prioridades, @"[^\d]", "").Trim();
                string saudacao = GetSaudacao();
                string nomeUsuario = Environment.UserName;

                //ConfirmarOT(PrioridadeSomenteNumeros);

                nomeUsuario = nomeUsuario.Replace('.', ' ');
                nomeUsuario = string.Join(" ", nomeUsuario.Split(' ').Select(p => char.ToUpper(p[0]) + p.Substring(1).ToLower()));

                string imagemOfelizFilePath = @"\\marconi\COMPANY SHARED FOLDER\OFELIZ\OFM\2.AN\2.CM\DP\4 Produção\Desenvolvimentos\Ficheiros Temp tekla artigos (Nao Apagar)\ofeliz_logo.png";


                MailMessage mail = new MailMessage();
                SmtpClient SmtpServer = new SmtpClient();
                mail.To.Add("clara.martins@ofeliz.com");
                mail.From = new MailAddress("alertas.cm@ofeliz.com");
                mail.Subject = $"{NumeroObra}-{Nomeobra}-" + "Processo de Soldadura";
                mail.IsBodyHtml = true;
                mail.Body = "<html><body><font face='Calibri' size='3'>" +
                           "<p>" + saudacao + "</p>" +
                           "<p>Venho por este meio informar, que já foi emitido dentro da pasta da obra em assunto, o Processo de Soldadura da <b> fase </b>" + PrioridadeSomenteNumeros + " da <b> Obra: </b>" + NumeroObra + "</p>" +
                           "</font>" +

                       "<font face='Calibri' size='3'><p>Melhores Cumprimentos,</p></font><br>" +
                       "<font face='Calibri' size='3'><b>" + nomeUsuario + "</b></font><br>" +
                       "<font face='Calibri' size='3'>Construção Metálica | Preparador</font><br>" +
                       "<font face='Calibri' size='3'>T + 351 253 080 609 *</font><br>" +
                       "<font color='red' face='Calibri' size='3'>ofeliz.com</font><br>" +

                       "</body></html>";


                SmtpServer.Host = "mx.ofeliz.com";
                SmtpServer.Port = 25;
                SmtpServer.UseDefaultCredentials = true;
                SmtpServer.EnableSsl = false;
                SmtpServer.DeliveryMethod = System.Net.Mail.SmtpDeliveryMethod.Network;

                int tentativas = 0;
                bool enviado = false;
                while (tentativas < 3 && !enviado)
                {
                    try
                    {
                        SmtpServer.Send(mail);
                        //MessageBox.Show("Email enviado");
                        enviado = true;
                    }
                    catch (SmtpException ex)
                    {
                        tentativas++;
                        if (tentativas >= 3)
                        {
                            MessageBox.Show("Erro ao enviar o e-mail. Informe a Clara que ja foi emitido o Processo de Soldadura.");

                            break;
                        }
                        else
                        {
                            MessageBox.Show($"Erro ao enviar o e-mail. Tentando novamente ({tentativas}/3).");
                            System.Threading.Thread.Sleep(3000);
                        }
                    }
                }
            }
            else
            { }
        }

        private void VerificarTarefaExistente()
        {
            string numerodaObra;

            if (!string.IsNullOrEmpty(TextBoxNObra.Text))
            {
                numerodaObra = TextBoxNObra.Text;
            }
            else
            {
                numerodaObra = ComboBoxObrasAdd.SelectedItem?.ToString();
            }

            string nomeObra = labelNomeObra.Text;
            string preparador = ComboBoxPreparadorAdd.SelectedItem?.ToString();
            string prioridades = ComboBoxPrioAdd.SelectedItem?.ToString();

            string query = @"
                            SELECT [Numero da Obra], [Nome da Obra], Tarefa, Preparador, Estado, Prioridades
                            FROM dbo.RegistoTarefas
                            WHERE [Numero da Obra] = @NumeroObra
                              AND [Nome da Obra] = @NomeObra
                              AND Preparador = @Preparador
                              AND Prioridades = @Prioridades
                              AND (Estado = 'Aguarda aprovação' OR Estado = 'Pendente')
                              AND Concluido = 0";

            ComunicaBD BD = new ComunicaBD();

            try
            {
                BD.ConectarBD();

                using (SqlCommand cmd = new SqlCommand(query, BD.GetConnection()))
                {
                    cmd.Parameters.AddWithValue("@NumeroObra", numerodaObra);
                    cmd.Parameters.AddWithValue("@NomeObra", nomeObra);
                    cmd.Parameters.AddWithValue("@Preparador", preparador);
                    cmd.Parameters.AddWithValue("@Prioridades", prioridades);

                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        string resultado = "Já existem tarefas com estas características:\n\n";

                        bool encontrouTarefas = false;

                        while (reader.Read())
                        {
                            encontrouTarefas = true;

                            //MessageBox.Show(
                            //    $"DEBUG:\n" +
                            //    $"Numero da Obra: {reader["Numero da Obra"]}\n" +
                            //    $"Nome da Obra: {reader["Nome da Obra"]}\n" +
                            //    $"Tarefa: {reader["Tarefa"]}\n" +
                            //    $"Preparador: {reader["Preparador"]}\n" +
                            //    $"Estado: {reader["Estado"]}\n" +
                            //    $"Prioridades: {reader["Prioridades"]}"                            
                            //);

                            resultado += $"Nome da Obra: {reader["Nome da Obra"]}\n" +
                                         $"Tarefa: {reader["Tarefa"]}\n" +
                                         $"Preparador: {reader["Preparador"]}\n" +
                                         $"Estado: {reader["Estado"]}\n" +
                                         $"Prioridades: {reader["Prioridades"]}\n\n";
                        }

                        if (encontrouTarefas)
                        {
                            MessageBox.Show(resultado, "Tarefas semelhantes já existentes", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }
                        else
                        {
                            MessageBox.Show("Nenhuma tarefa semelhante foi encontrada.", "Verificação Concluída", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erro ao verificar tarefa: " + ex.Message);
            }
            finally
            {
                BD.DesonectarBD();
            }
        }

        private bool ExisteTarefaEmAndamento()
        {
            string Numeroobra = null;
            string NomeObra = null;
            string Tarefa = null;
            string Preparador = null;
            string Prioridades = null;
            string DataTarefa = null;
            string horafinal = "00:00:00";
            string QtdHora = "00:00:00";

            if (DataGridViewAddTarefas.SelectedRows.Count > 0)
            {
                Numeroobra = DataGridViewAddTarefas.SelectedRows[0].Cells["Numero da Obra"].Value?.ToString().Trim();
                Tarefa = DataGridViewAddTarefas.SelectedRows[0].Cells["Tarefa"].Value?.ToString().Trim();
                Preparador = DataGridViewAddTarefas.SelectedRows[0].Cells["Preparador"].Value?.ToString().Trim();
                Prioridades = DataGridViewAddTarefas.SelectedRows[0].Cells["Prioridades"].Value?.ToString().Trim();
                DataTarefa = DateTime.Now.ToString("dd/MM/yyyy"); 
            }
            else if (DataGridViewPendente.SelectedRows.Count > 0)
            {
                Numeroobra = DataGridViewPendente.SelectedRows[0].Cells["Numero da Obra"].Value?.ToString().Trim();
                Tarefa = DataGridViewPendente.SelectedRows[0].Cells["Tarefa"].Value?.ToString().Trim();
                Preparador = DataGridViewPendente.SelectedRows[0].Cells["Preparador"].Value?.ToString().Trim();
                Prioridades = DataGridViewPendente.SelectedRows[0].Cells["Prioridades"].Value?.ToString().Trim();
                DataTarefa = DateTime.Now.ToString("dd/MM/yyyy");
            }
            else if (DataGridViewAguardarAprovação.SelectedRows.Count > 0)
            {
                Numeroobra = DataGridViewAguardarAprovação.SelectedRows[0].Cells["Numero da Obra"].Value?.ToString().Trim();
                Tarefa = DataGridViewAguardarAprovação.SelectedRows[0].Cells["Tarefa"].Value?.ToString().Trim();
                Preparador = DataGridViewAguardarAprovação.SelectedRows[0].Cells["Preparador"].Value?.ToString().Trim();
                Prioridades = DataGridViewAguardarAprovação.SelectedRows[0].Cells["Prioridades"].Value?.ToString().Trim();
                DataTarefa = DateTime.Now.ToString("dd/MM/yyyy");
            }

            ComunicaBD comunicaBD = new ComunicaBD();
            try
            {
                comunicaBD.ConectarBD();
                              
                string query = @"SELECT 1 FROM dbo.RegistoTempo 
                         WHERE [Numero da Obra] = @NumeroObra  
                            AND Tarefa = @Tarefa
                            AND Preparador = @Preparador
                            AND Prioridade = @Prioridade
                            AND [Data da Tarefa] = @DataTarefa
                            AND [Qtd de Hora] = @QtdHora 
                            AND [Hora Final] = @Horafinal";

               

                using (SqlCommand cmd = new SqlCommand(query, comunicaBD.GetConnection()))
                {
                    cmd.Parameters.AddWithValue("@NumeroObra", Numeroobra);
                    cmd.Parameters.AddWithValue("@Tarefa", Tarefa);
                    cmd.Parameters.AddWithValue("@Preparador", Preparador);
                    cmd.Parameters.AddWithValue("@Prioridade", Prioridades);
                    cmd.Parameters.AddWithValue("@DataTarefa", DataTarefa);
                    cmd.Parameters.AddWithValue("@QtdHora", QtdHora);
                    cmd.Parameters.AddWithValue("@Horafinal", horafinal);

                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        if (reader.HasRows)
                        {
                            MessageBox.Show(
                                            $"O Preparador: {Preparador} tem a tarefa a decorrer",
                                            "Aviso",
                                            MessageBoxButtons.OK,
                                            MessageBoxIcon.Warning
                                        );
                            return true;
                        }
                        else
                        {                            
                            return false;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erro ao verificar tarefa em andamento: " + ex.Message);
                return true; 
            }
            finally
            {
                comunicaBD.DesonectarBD();
            }
        }
    }
}









