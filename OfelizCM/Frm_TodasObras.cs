using GMap.NET;
using GMap.NET.MapProviders;
using GMap.NET.WindowsForms;
using GMap.NET.WindowsForms.Markers;
using GMap.NET.WindowsPresentation;
using LiveCharts;
using LiveCharts.Wpf;
using MaterialSkin;
using MaterialSkin.Controls;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using ServiceStack.Script;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;
using System.Windows.Media;
using static ServiceStack.Script.Lisp;
using Color = System.Drawing.Color;
using Cor = System.Windows.Media;

namespace OfelizCM
{
    public partial class Frm_TodasObras : Form
    {
        private ChamarX _hook;
        private Image imgOriginal;
        private Image imgGif;
        public Frm_TodasObras()
        {
            InitializeComponent();
            _hook = new ChamarX();
        }

        public class ChamarX
        {
            public bool IsActive { get; private set; }
            public void Start()
            {
                IsActive = true;
            }
            public void Stop()
            {
                IsActive = false;
            }
        }
        private void Frm_TodasObras_Load(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Maximized;
            PanelTabelas.Location = new Point(215, 60);
            PanelTabelas.Size = new Size(2342, 1395);
            this.totalObrasTableAdapter.Fill(this.tempoPreparacaoDataSet.TotalObras);
            ComunicarTabelasSemAnofecho();
            AtualizarDados();
            CalcularTabelas();
            CarregarGraficos();
            Calcular();
            Limpar(DataGridViewOrcamentacaoObras);

            ConfigurarComboBoxTipologia();
            ConfigurarComboBoxPreparadorResponsavel();
            CarregarPreparadoresNaComboBox();
            DataGridViewOrcamentacaoObras.CellBeginEdit += DataGridViewOrcamentacaoObras_CellBeginEdit;
            DataGridViewOrcamentacaoObras.CellEndEdit += DataGridViewOrcamentacaoObras_CellEndEdit;
            DataGridViewRealObras.SelectionChanged += DataGridViewRealObras_SelectionChanged;
            DataGridViewOrcamentacaoObras.SelectionChanged += DataGridViewRealObrasNome_SelectionChanged;
            guna2VScrollBar1.Scroll += Guna2VScrollBar_Scroll;
            DataGridViewRealObras.Scroll += DataGridViewRealObras_Scroll;
            DataGridViewOrcamentacaoObras.Scroll += DataGridViewOrcamentacaoObras_Scroll;
            DataGridViewConclusaoObras.Scroll += DataGridViewConclusaoObras_Scroll;
            VerificarUsuario();
            InicializarSincronizacaoDeRolagem();
            CarregarPastasNaComboBoxAno();
            AdicionarSufixosNasColunas();
            CarregarTipologiaNaComboBox();
            Modificartabelas();
            if (!string.IsNullOrEmpty(Properties.Settings.Default.PrecoOrcamentado))
            {
                LabelPrecoOrcamentado.Text = Properties.Settings.Default.PrecoOrcamentado;
            }
            imgOriginal = ButtonLimaprtaabela.Image;
            imgGif = Properties.Resources.refresh;
            ButtonLimaprtaabela.Image = imgOriginal;
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
                    bool autorizado = false;

                    if (result != null)
                        autorizado = Convert.ToBoolean(result);

                    string nomeUsuario2 = Properties.Settings.Default.NomeUsuario;
                    if (nomeUsuario2 == "ofelizcmadmin" || nomeUsuario2 == "helder.silva")
                        autorizado = true;
                    DefinirVisibilidadeBotoes(autorizado);
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
        private void DefinirVisibilidadeBotoes(bool visivel)
        {
            PanelInserirAtualizar.Visible = visivel;
            PanelExportar.Visible = visivel;
        }
        private void ComunicarTabelas()
        {            
            VisualizarTabelaOrcamentacao();
            VisualizarTabelaReal();
            VisualizarTabelaConcluido();       
        }
        private void ComunicarTabelasSemAnofecho()
        {
            VisualizarTabelaOrcamentacaoSemAnofecho();
            VisualizarTabelaRealSemAnofecho();
            VisualizarTabelaConcluidoSemAnofecho();
        }
        private void AtualizarDados()
        {
            AtualizarTabelaRealnaBd();
            AtualizarTabelaRealnaBd2();
            AtualizarTabelaRealnaBd3();
            AtualizarTabelaConclusaoBD();
            AtualizarTabelaConclusaoBD2();
        }
        private void CalcularTabelas()
        {
            CalcularTotaisEInserirNaBaseDeDados();
            CalcularTotaisReal();
            CalcularTotaisRealPercentagem();
            CalcularTotaisConclusao();
        }
        private void CarregarGraficos()
        {
            VisualizarGraficoTotalHoras();
            VisualizarGraficoObrasHoras();
            VisualizarGraficoObrasValor();
            VisualizarGraficoObrasPercentagem();
            VisualizarGraficoPiePercentagem();
        }
        private void FiltrarTipologia()
        {
            string Tipologia = ComboBoxTipologiaFiltro.SelectedItem.ToString();
            ComunicaBD BD = new ComunicaBD();
            {
                try
                {
                    BD.ConectarBD();
                    string queryorçamentacao = "SELECT * FROM dbo.Orçamentação WHERE (Tipologia) = @Tipologia";
                    string queryreal = "SELECT * FROM dbo.RealObras WHERE (Tipologia) = @Tipologia";
                    string queryconclusao = "SELECT * FROM dbo.ConclusaoObras WHERE (Tipologia) = @Tipologia";
                    using (var command = new SqlCommand(queryorçamentacao, BD.GetConnection()))
                    {
                        command.Parameters.AddWithValue("@Tipologia", Tipologia);
                        DataTable dataTable = new DataTable();
                        using (var adapter = new SqlDataAdapter(command))
                        {
                            adapter.Fill(dataTable);
                        }
                        DataGridViewOrcamentacaoObras.DataSource = dataTable;
                        ConfigurarColunasOrcamentacao();
                        DataGridViewOrcamentacaoObras.ClearSelection();
                    }
                    using (var command = new SqlCommand(queryreal, BD.GetConnection()))
                    {
                        command.Parameters.AddWithValue("@Tipologia", Tipologia);

                        DataTable dataTable = new DataTable();

                        using (var adapter = new SqlDataAdapter(command))
                        {
                            adapter.Fill(dataTable);
                        }
                        DataGridViewRealObras.DataSource = dataTable;
                        DataGridViewRealObras.Columns["Nome da Obra"].Visible = false;
                        DataGridViewRealObras.Columns["Preparador Responsavel"].Visible = false;
                        DataGridViewRealObras.ClearSelection();
                    }
                    using (var command = new SqlCommand(queryconclusao, BD.GetConnection()))
                    {
                        command.Parameters.AddWithValue("@Tipologia", Tipologia);

                        DataTable dataTable = new DataTable();

                        using (var adapter = new SqlDataAdapter(command))
                        {
                            adapter.Fill(dataTable);
                        }
                        DataGridViewConclusaoObras.DataSource = dataTable;
                        DataGridViewConclusaoObras.Columns["Nome da Obra"].Visible = false;
                        DataGridViewConclusaoObras.Columns["Preparador Responsavel"].Visible = false;
                        DataGridViewConclusaoObras.Columns["Total Horas"].Width = 90;
                        DataGridViewConclusaoObras.Columns["Total Valor"].Width = 90;
                        DataGridViewConclusaoObras.Columns["Percentagem Total"].Width = 70;
                        DataGridViewConclusaoObras.Columns["Dias de Preparação"].Width = 70;
                        DataGridViewConclusaoObras.ClearSelection();
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Erro ao conectar à base de dados: " + ex.Message);
                }
            }
        }
        private void FiltrarPorAno()
        {
            string anoSelecionado = ComboBoxAnoAdd.SelectedItem.ToString();
            ComunicaBD BD = new ComunicaBD();
            {
                try
                {
                    BD.ConectarBD();
                    string queryorçamentacao = "SELECT * FROM dbo.Orçamentação WHERE YEAR([Ano de fecho]) = @AnoFecho";
                    string queryreal = "SELECT * FROM dbo.RealObras WHERE YEAR([Ano de fecho]) = @AnoFecho";
                    string queryconclusao = "SELECT * FROM dbo.ConclusaoObras WHERE YEAR([Ano de fecho]) = @AnoFecho";

                    using (var command = new SqlCommand(queryorçamentacao, BD.GetConnection()))
                    {
                        command.Parameters.AddWithValue("@AnoFecho", anoSelecionado);

                        DataTable dataTable = new DataTable();

                        using (var adapter = new SqlDataAdapter(command))
                        {
                            adapter.Fill(dataTable);
                        }
                        DataGridViewOrcamentacaoObras.DataSource = dataTable;
                        ConfigurarColunasOrcamentacao();
                        DataGridViewOrcamentacaoObras.ClearSelection();
                    }
                    using (var command = new SqlCommand(queryreal, BD.GetConnection()))
                    {
                        command.Parameters.AddWithValue("@AnoFecho", anoSelecionado);
                        DataTable dataTable = new DataTable();
                        using (var adapter = new SqlDataAdapter(command))
                        {
                            adapter.Fill(dataTable);
                        }
                        DataGridViewRealObras.DataSource = dataTable;
                        DataGridViewRealObras.Columns["Nome da Obra"].Visible = false;
                        DataGridViewRealObras.Columns["Preparador Responsavel"].Visible = false;
                        DataGridViewRealObras.ClearSelection();
                    }
                    using (var command = new SqlCommand(queryconclusao, BD.GetConnection()))
                    {
                        command.Parameters.AddWithValue("@AnoFecho", anoSelecionado);
                        DataTable dataTable = new DataTable();
                        using (var adapter = new SqlDataAdapter(command))
                        {
                            adapter.Fill(dataTable);
                        }
                        DataGridViewConclusaoObras.DataSource = dataTable;
                        DataGridViewConclusaoObras.Columns["Nome da Obra"].Visible = false;
                        DataGridViewConclusaoObras.Columns["Preparador Responsavel"].Visible = false;
                        DataGridViewConclusaoObras.Columns["Total Horas"].Width = 90;
                        DataGridViewConclusaoObras.Columns["Total Valor"].Width = 90;
                        DataGridViewConclusaoObras.Columns["Percentagem Total"].Width = 70;
                        DataGridViewConclusaoObras.Columns["Dias de Preparação"].Width = 70;
                        DataGridViewConclusaoObras.ClearSelection();
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Erro ao conectar à base de dados: " + ex.Message);
                }
            }
        }
        private void InserirAnofecho()
        {
            if (DataGridViewOrcamentacaoObras.SelectedRows.Count > 0)
            {
                string NumeroObra = DataGridViewOrcamentacaoObras.SelectedRows[0].Cells["Numero da Obra"].Value.ToString();
                string anoSelecionado = ComboBoxAnoAdd2.SelectedItem?.ToString();
                if (anoSelecionado == null)
                {
                    MessageBox.Show("Selecione um ano.");
                    return;
                }
                ComunicaBD BD = new ComunicaBD();
                try
                {
                    BD.ConectarBD();

                    string queryorcamentacao = "UPDATE dbo.Orçamentação SET [Ano de fecho] = @AnoFecho " +
                                               "WHERE [Numero da Obra] = @NumeroObra";

                    string queryreal = "UPDATE dbo.RealObras SET [Ano de fecho] = @AnoFecho " +
                                       "WHERE [Numero da Obra] = @NumeroObra";

                    string queryconclusao = "UPDATE dbo.ConclusaoObras SET [Ano de fecho] = @AnoFecho " +
                                            "WHERE [Numero da Obra] = @NumeroObra";

                    using (var command = new SqlCommand(queryorcamentacao, BD.GetConnection()))
                    {
                        command.Parameters.AddWithValue("@AnoFecho", anoSelecionado);
                        command.Parameters.AddWithValue("@NumeroObra", NumeroObra);
                        command.ExecuteNonQuery();
                    }
                    using (var command = new SqlCommand(queryreal, BD.GetConnection()))
                    {
                        command.Parameters.AddWithValue("@AnoFecho", anoSelecionado);
                        command.Parameters.AddWithValue("@NumeroObra", NumeroObra);
                        command.ExecuteNonQuery();
                    }
                    using (var command = new SqlCommand(queryconclusao, BD.GetConnection()))
                    {
                        command.Parameters.AddWithValue("@AnoFecho", anoSelecionado);
                        command.Parameters.AddWithValue("@NumeroObra", NumeroObra);
                        command.ExecuteNonQuery();
                    }
                    DataGridViewOrcamentacaoObras.ClearSelection();
                    DataGridViewRealObras.ClearSelection();
                    DataGridViewConclusaoObras.ClearSelection();
                    MessageBox.Show("Ano de fecho Inserido com Sucesso");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Erro ao inserir o Ano de fecho na base de dados: " + ex.Message);
                }
                finally
                {
                    BD.DesonectarBD();
                }
            }
            else
            {
                MessageBox.Show("Nenhuma obra selecionada.");
            }
        }
        private void LimparAnoFecho()
        {
            if (DataGridViewOrcamentacaoObras.SelectedRows.Count > 0)
            {
                string NumeroObra = DataGridViewOrcamentacaoObras.SelectedRows[0].Cells["Numero da Obra"].Value.ToString();
                ComunicaBD BD = new ComunicaBD();
                try
                {
                    BD.ConectarBD();

                    string queryorcamentacao = "UPDATE dbo.Orçamentação SET [Ano de fecho] = ' ' " +
                                               "WHERE [Numero da Obra] = @NumeroObra";

                    string queryreal = "UPDATE dbo.RealObras SET [Ano de fecho] =  ' ' " +
                                       "WHERE [Numero da Obra] = @NumeroObra";

                    string queryconclusao = "UPDATE dbo.ConclusaoObras SET [Ano de fecho] = ' ' " +
                                            "WHERE [Numero da Obra] = @NumeroObra";

                    using (var command = new SqlCommand(queryorcamentacao, BD.GetConnection()))
                    {
                        command.Parameters.AddWithValue("@NumeroObra", NumeroObra);
                        command.ExecuteNonQuery();
                    }

                    using (var command = new SqlCommand(queryreal, BD.GetConnection()))
                    {
                        command.Parameters.AddWithValue("@NumeroObra", NumeroObra);
                        command.ExecuteNonQuery();
                    }

                    using (var command = new SqlCommand(queryconclusao, BD.GetConnection()))
                    {
                        command.Parameters.AddWithValue("@NumeroObra", NumeroObra);
                        command.ExecuteNonQuery();
                    }
                    DataGridViewOrcamentacaoObras.ClearSelection();
                    DataGridViewRealObras.ClearSelection();
                    DataGridViewConclusaoObras.ClearSelection();

                    MessageBox.Show("Ano de fecho removido com sucesso.");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Erro ao remover o Ano de fecho na base de dados: " + ex.Message);
                }
                finally
                {
                    BD.DesonectarBD();
                }
            }
            else
            {
                MessageBox.Show("Nenhuma obra selecionada.");
            }
        }
        private void InserirTipologia()
        {
            if (DataGridViewOrcamentacaoObras.SelectedRows.Count > 0)
            {
                string NumeroObra = DataGridViewOrcamentacaoObras.SelectedRows[0].Cells["Numero da Obra"].Value.ToString();
                string Tipologia = ComboBoxTipologiaInserir.SelectedItem?.ToString();

                if (Tipologia == null)
                {
                    MessageBox.Show("Selecione uma Tipologia");
                    return;
                }
                ComunicaBD BD = new ComunicaBD();
                try
                {
                    BD.ConectarBD();
                    string queryorcamentacao = "UPDATE dbo.Orçamentação SET Tipologia = @Tipologia " +
                                               "WHERE [Numero da Obra] = @NumeroObra";
                    string queryreal = "UPDATE dbo.RealObras SET Tipologia = @Tipologia " +
                                       "WHERE [Numero da Obra] = @NumeroObra";
                    string queryconclusao = "UPDATE dbo.ConclusaoObras SET Tipologia = @Tipologia " +
                                            "WHERE [Numero da Obra] = @NumeroObra";

                    using (var command = new SqlCommand(queryorcamentacao, BD.GetConnection()))
                    {
                        command.Parameters.AddWithValue("@Tipologia", Tipologia);
                        command.Parameters.AddWithValue("@NumeroObra", NumeroObra);
                        command.ExecuteNonQuery();
                    }
                    using (var command = new SqlCommand(queryreal, BD.GetConnection()))
                    {
                        command.Parameters.AddWithValue("@Tipologia", Tipologia);
                        command.Parameters.AddWithValue("@NumeroObra", NumeroObra);
                        command.ExecuteNonQuery();
                    }
                    using (var command = new SqlCommand(queryconclusao, BD.GetConnection()))
                    {
                        command.Parameters.AddWithValue("@Tipologia", Tipologia);
                        command.Parameters.AddWithValue("@NumeroObra", NumeroObra);
                        command.ExecuteNonQuery();
                    }
                    DataGridViewOrcamentacaoObras.ClearSelection();
                    DataGridViewRealObras.ClearSelection();
                    DataGridViewConclusaoObras.ClearSelection();
                    MessageBox.Show("Tipologia da obra Inserida com Sucesso");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Erro ao inserir a Tipologia da obra na base de dados: " + ex.Message);
                }
                finally
                {
                    BD.DesonectarBD();
                }
            }
            else
            {
                MessageBox.Show("Tipologia selecionada.");
            }
        }       
        private void MudarTabelaGrafico_CheckedChanged(object sender, EventArgs e)
        {
            if (!_hook.IsActive)
            {
                _hook.Start();
                PanelTabelas.Visible = false;
                labeltabelagraficos.Text = "Gráficos Controlo das Obra";
                labelgraficostabela.Text = "Gráficos";
                CarregarGraficos();
                PanelTotalObras.Visible = true;
                PanelChartControlo.Visible = true;
                PanelchartTotalHoras.Visible = true;
                PanelchartObrasHoras.Visible = true;
                PanelchartTotalValorObras.Visible = true;
            }
            else
            {
                _hook.Stop();
                PanelTabelas.Visible = true;
                labeltabelagraficos.Text = "Tabela Controlo das Obra";
                labelgraficostabela.Text = "Tabela";
                PanelTotalObras.Visible = false;
                PanelChartControlo.Visible = false;
                PanelchartTotalHoras.Visible = false;
                PanelchartObrasHoras.Visible = false;
                PanelchartTotalValorObras.Visible = false;
            }
        }
        private void VisualizarTabelaOrcamentacao()
        {
            var tabela = new Mostartabelas();
            DataTable dataTable = tabela.TabelaOrcamentacao();

            DataGridViewOrcamentacaoObras.DataSource = dataTable;
            DataGridViewOrcamentacaoObras.ClearSelection();
            DataGridViewOrcamentacaoObras.ScrollBars = ScrollBars.Horizontal;
            DataGridViewOrcamentacaoObras.Columns["Id"].Visible = false;
            ConfigurarColunasOrcamentacao();
        }
        private void VisualizarTabelaReal()
        {
            var tabela = new Mostartabelas();
            DataTable dataTable = tabela.TabelaReal();

            DataGridViewRealObras.DataSource = dataTable;
            DataGridViewRealObras.ClearSelection();
            DataGridViewRealObras.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
            DataGridViewRealObras.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
            DataGridViewRealObras.Columns["Id"].Visible = false;
            DataGridViewRealObras.Columns["Numero da Obra"].Visible = false;
            DataGridViewRealObras.Columns["Ano de fecho"].Visible = false;
            DataGridViewRealObras.Columns["Tipologia"].Visible = false;
            DataGridViewRealObras.ReadOnly = true;
            DataGridViewRealObras.Columns["KG Estrutura"].Width = 90;
            DataGridViewRealObras.Columns["KG/Euro Estrutura"].HeaderText = "Euro/kg Estrutura";
            DataGridViewRealObras.ScrollBars = ScrollBars.Horizontal;
            int linhaInicial = 25;
            if (linhaInicial < DataGridViewRealObras.Rows.Count)
                DataGridViewRealObras.FirstDisplayedScrollingRowIndex = linhaInicial;
        }
        private void VisualizarTabelaConcluido()
        {
            var tabela = new Mostartabelas();
            DataTable dataTable = tabela.TabelaConclusao();

            DataGridViewConclusaoObras.DataSource = dataTable;
            DataGridViewConclusaoObras.ClearSelection();
            DataGridViewConclusaoObras.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            DataGridViewConclusaoObras.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            DataGridViewConclusaoObras.Columns["Id"].Visible = false;
            DataGridViewConclusaoObras.Columns["Numero da Obra"].Visible = false;
            DataGridViewConclusaoObras.Columns["Ano de fecho"].Visible = false;
            DataGridViewConclusaoObras.Columns["Tipologia"].Visible = false;
            DataGridViewConclusaoObras.ReadOnly = true;
            DataGridViewConclusaoObras.Columns["Total Horas"].Width = 90;
            DataGridViewConclusaoObras.Columns["Total Valor"].Width = 90;
            DataGridViewConclusaoObras.Columns["Percentagem Total"].Width = 70;
            DataGridViewConclusaoObras.Columns["Dias de Preparação"].Width = 70;
        }
        private void VisualizarTabelaOrcamentacaoSemAnofecho()
        {
            var tabela = new Mostartabelassemanofecho();
            DataTable dataTable = tabela.TabelaOrcamentacao();

            DataGridViewOrcamentacaoObras.DataSource = dataTable;
            DataGridViewOrcamentacaoObras.ClearSelection();
            DataGridViewOrcamentacaoObras.ScrollBars = ScrollBars.Horizontal;
            DataGridViewOrcamentacaoObras.Columns["Id"].Visible = false;
            ConfigurarColunasOrcamentacao();
        }
        private void VisualizarTabelaRealSemAnofecho()
        {
            var tabela = new Mostartabelassemanofecho();
            DataTable dataTable = tabela.TabelaReal();

            DataGridViewRealObras.DataSource = dataTable;
            DataGridViewRealObras.ClearSelection();
            DataGridViewRealObras.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
            DataGridViewRealObras.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
            DataGridViewRealObras.Columns["Id"].Visible = false;
            DataGridViewRealObras.Columns["Numero da Obra"].Visible = false;
            DataGridViewRealObras.Columns["Ano de fecho"].Visible = false;
            DataGridViewRealObras.Columns["Tipologia"].Visible = false;
            DataGridViewRealObras.ReadOnly = true;
            DataGridViewRealObras.Columns["KG Estrutura"].Width = 90;
            DataGridViewRealObras.Columns["KG/Euro Estrutura"].HeaderText = "Euro/kg Estrutura";
            DataGridViewRealObras.ScrollBars = ScrollBars.Horizontal;
            int linhaInicial = 25;
            if (linhaInicial < DataGridViewRealObras.Rows.Count)
                DataGridViewRealObras.FirstDisplayedScrollingRowIndex = linhaInicial;
        }
        private void VisualizarTabelaConcluidoSemAnofecho()
        {
            var tabela = new Mostartabelassemanofecho();
            DataTable dataTable = tabela.TabelaConclusao();

            DataGridViewConclusaoObras.DataSource = dataTable;
            DataGridViewConclusaoObras.ClearSelection();
            DataGridViewConclusaoObras.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            DataGridViewConclusaoObras.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            DataGridViewConclusaoObras.Columns["Id"].Visible = false;
            DataGridViewConclusaoObras.Columns["Numero da Obra"].Visible = false;
            DataGridViewConclusaoObras.Columns["Ano de fecho"].Visible = false;
            DataGridViewConclusaoObras.Columns["Tipologia"].Visible = false;
            DataGridViewConclusaoObras.ReadOnly = true;
            DataGridViewConclusaoObras.Columns["Total Horas"].Width = 90;
            DataGridViewConclusaoObras.Columns["Total Valor"].Width = 90;
            DataGridViewConclusaoObras.Columns["Percentagem Total"].Width = 70;
            DataGridViewConclusaoObras.Columns["Dias de Preparação"].Width = 70;
        }
        public void VisualizarGraficoPiePercentagem()
        {
            MostarGraficos grafico = new MostarGraficos();

            var pieChart = pieChartControlo.Child as LiveCharts.Wpf.PieChart;

            if (pieChart != null)
            {
                pieChart.Series = grafico.CarregarGraficoRedondo();
                pieChart.LegendLocation = LegendLocation.Top;
            }
        }
        public void VisualizarGraficoPiePercentagemTipologiaeAno()
        {
            var grafico = new MostarGraficos();
            DataTable tabela = DataGridViewRealObras.DataSource as DataTable;

            var pieChart = pieChartControlo.Child as LiveCharts.Wpf.PieChart;

            if (pieChart != null)
            {
                pieChart.Series = grafico.CalcularPercentagemTipologia(tabela);
                pieChart.LegendLocation = LiveCharts.LegendLocation.Top;
            }
        }
        private void VisualizarGraficoObrasPercentagem()
        {
            var chartWpf = ChartTotalObrasPercentagem.Child as LiveCharts.Wpf.CartesianChart;
            if (chartWpf == null) return;
            var service = new MostarGraficos();
            var (series, labels) = service.CarregarPercentagemTodasObras();

            chartWpf.Series = series;
            foreach (var s in chartWpf.Series.OfType<ColumnSeries>())
            {
                s.MaxColumnWidth = 30;
                s.ColumnPadding = 10;
                s.DataLabels = true;
                s.LabelPoint = point => $"{point.Y:N1} %";
            }

            chartWpf.AxisX.Clear();
            chartWpf.AxisX.Add(new LiveCharts.Wpf.Axis
            {
                Title = "Número da Obra",
                FontSize = 15,
                Labels = labels,
                LabelsRotation = 45,
                Separator = new LiveCharts.Wpf.Separator { Step = 1 },
            });

            chartWpf.AxisY.Clear();
            chartWpf.AxisY.Add(new LiveCharts.Wpf.Axis
            {
                Title = "Percentagem (%)",
                LabelFormatter = value => value.ToString("N1") + " %",
                Sections = new SectionsCollection
             {
            new AxisSection
             {
                Value = 100,
                Stroke = Cor.Brushes.Red,
                StrokeThickness = 1,
                StrokeDashArray = new DoubleCollection { 4 }
             }
            }
            });
            chartWpf.Hoverable = true;
            chartWpf.DataTooltip = new DefaultTooltip();
            chartWpf.DataHover += (sender, chartPoint) =>
            {
                int index = (int)chartPoint.X;
                string numeroObra = labels[index];
                string nomeObra = new MostarGraficos().ObterNomeObra(numeroObra);
                double percentagem = chartPoint.Y;

                labelobra.Text = $"{nomeObra} | {percentagem:N1} %";
            };
        }
        private void VisualizarGraficoTotalHoras()
        {
            var chartWpf = chartTotalHoras.Child as LiveCharts.Wpf.CartesianChart;
            if (chartWpf == null) return;

            var service = new MostarGraficos();
            var (series, labels) = service.CarregarTotalTodasObras();

            chartWpf.Series = series;

            foreach (var s in chartWpf.Series.OfType<ColumnSeries>())
            {
                s.MaxColumnWidth = 30;
                s.ColumnPadding = 10;
            }

            chartWpf.AxisX.Clear();
            chartWpf.AxisX.Add(new LiveCharts.Wpf.Axis
            {
                Title = "Obra",
                FontSize = 15,
                Labels = labels,
                Separator = new LiveCharts.Wpf.Separator { Step = 1 },

                LabelsRotation = 45
            });

            chartWpf.AxisY.Clear();
            chartWpf.AxisY.Add(new LiveCharts.Wpf.Axis
            {
                Title = "Horas",
                FontSize = 15,
                LabelFormatter = value => value + "h"
            });

            chartWpf.LegendLocation = LegendLocation.Top;
        }
        private void VisualizarGraficoObrasHoras()
        {
            var chartWpf = chartObrasHoras.Child as LiveCharts.Wpf.CartesianChart;
            if (chartWpf == null) return;
            var service = new MostarGraficos();
            var (series, labels) = service.CarregarGraficoObrasHoras();
            chartWpf.Series = series;
            foreach (var s in chartWpf.Series.OfType<ColumnSeries>())
            {
                s.MaxColumnWidth = 30;
                s.ColumnPadding = 10;
            }
            chartWpf.AxisX.Clear();
            chartWpf.AxisX.Add(new LiveCharts.Wpf.Axis
            {
                Title = "Número da Obra",
                FontSize = 15,
                Labels = labels,
                Separator = new LiveCharts.Wpf.Separator { Step = 1 },
                LabelsRotation = 45
            });

            chartWpf.AxisY.Clear();
            chartWpf.AxisY.Add(new LiveCharts.Wpf.Axis
            {
                Title = "Horas",
                FontSize = 15,
                LabelFormatter = value => value + " h"
            });
            chartWpf.LegendLocation = LegendLocation.Top;
            chartWpf.Hoverable = true;
            chartWpf.DataTooltip = new DefaultTooltip();
            chartWpf.DataHover += (sender, chartPoint) =>
            {
                int index = (int)chartPoint.X;
                string numeroObra = labels[index];
                string nomeObra = new MostarGraficos().ObterNomeObra(numeroObra);
                double percentagem = chartPoint.Y;

                labelobrasHoras.Text = $"{nomeObra} | {percentagem:N1} h";
            };
        }
        private void VisualizarGraficoObrasValor()
        {
            var chartWpf = chartTotalValorObras.Child as LiveCharts.Wpf.CartesianChart;
            if (chartWpf == null) return;
            var service = new MostarGraficos();
            var (series, labels) = service.CarregarGraficoObrasValor();
            chartWpf.Series = series;
            foreach (var s in chartWpf.Series.OfType<ColumnSeries>())
            {
                s.MaxColumnWidth = 30;
                s.ColumnPadding = 10;
            }
            chartWpf.AxisX.Clear();
            chartWpf.AxisX.Add(new LiveCharts.Wpf.Axis
            {
                Title = "Número da Obra",
                FontSize = 15,
                Labels = labels,
                Separator = new LiveCharts.Wpf.Separator { Step = 1 },
                LabelsRotation = 45
            });
            chartWpf.AxisY.Clear();
            chartWpf.AxisY.Add(new LiveCharts.Wpf.Axis
            {
                Title = "Valor (€)",
                FontSize = 15,
                LabelFormatter = value => "€" + value.ToString("N0")
            });
            chartWpf.LegendLocation = LegendLocation.Top;
            chartWpf.Hoverable = true;
            chartWpf.DataTooltip = new DefaultTooltip();
            chartWpf.DataHover += (sender, chartPoint) =>
            {
                int index = (int)chartPoint.X;
                string numeroObra = labels[index];
                string nomeObra = new MostarGraficos().ObterNomeObra(numeroObra);
                double percentagem = chartPoint.Y;

                labelobraeuros.Text = $"{nomeObra} |  {percentagem:N2} €";
            };
        }
        private void VisualizarGraficoObrasValorTipologia()
        {
            string tipologia = ComboBoxTipologiaFiltro.SelectedItem?.ToString();
            string campofiltro = "Tipologia";
            if (string.IsNullOrEmpty(tipologia))
            {
                MessageBox.Show("Selecione uma Tipologia para filtrar o gráfico.");
                return;
            }
            var chartWpf = chartTotalValorObras.Child as LiveCharts.Wpf.CartesianChart;
            if (chartWpf == null) return;
            var service = new MostarGraficos();
            var (series, labels) = service.CarregarGraficoObrasValor(campofiltro, tipologia);

            chartWpf.Series = series;

            foreach (var s in chartWpf.Series.OfType<ColumnSeries>())
            {
                s.MaxColumnWidth = 30;
                s.ColumnPadding = 10;
            }
            chartWpf.AxisX.Clear();
            chartWpf.AxisX.Add(new LiveCharts.Wpf.Axis
            {
                Title = "Número da Obra",
                FontSize = 15,
                Labels = labels,
                Separator = new LiveCharts.Wpf.Separator { Step = 1 },
                LabelsRotation = 45
            });
            chartWpf.AxisY.Clear();
            chartWpf.AxisY.Add(new LiveCharts.Wpf.Axis
            {
                Title = "Valor (€)",
                FontSize = 15,
                LabelFormatter = value => "€" + value.ToString("N0")
            });

            chartWpf.LegendLocation = LegendLocation.Top;
            chartWpf.Hoverable = true;
            chartWpf.DataTooltip = new DefaultTooltip();
            chartWpf.DataHover += (sender, chartPoint) =>
            {
                int index = (int)chartPoint.X;
                string numeroObra = labels[index];
                string nomeObra = new MostarGraficos().ObterNomeObra(numeroObra);
                double percentagem = chartPoint.Y;

                labelobraeuros.Text = $"{nomeObra} |  {percentagem:N2} €";
            };
        }
        private void VisualizarGraficoObrasValorAno()
        {
            string anofecho = ComboBoxAnoAdd.SelectedItem?.ToString();
            string campofiltro = "[Ano de fecho]";
            if (string.IsNullOrEmpty(anofecho))
            {
                MessageBox.Show("Selecione uma Tipologia para filtrar o gráfico.");
                return;
            }
            var chartWpf = chartTotalValorObras.Child as LiveCharts.Wpf.CartesianChart;
            if (chartWpf == null) return;
            var service = new MostarGraficos();
            var (series, labels) = service.CarregarGraficoObrasValor(campofiltro, anofecho);

            chartWpf.Series = series;

            foreach (var s in chartWpf.Series.OfType<ColumnSeries>())
            {
                s.MaxColumnWidth = 30;
                s.ColumnPadding = 10;
            }
            chartWpf.AxisX.Clear();
            chartWpf.AxisX.Add(new LiveCharts.Wpf.Axis
            {
                Title = "Número da Obra",
                FontSize = 15,
                Labels = labels,
                Separator = new LiveCharts.Wpf.Separator { Step = 1 },
                LabelsRotation = 45
            });
            chartWpf.AxisY.Clear();
            chartWpf.AxisY.Add(new LiveCharts.Wpf.Axis
            {
                Title = "Valor (€)",
                FontSize = 15,
                LabelFormatter = value => "€" + value.ToString("N0")
            });

            chartWpf.LegendLocation = LegendLocation.Top;
            chartWpf.Hoverable = true;
            chartWpf.DataTooltip = new DefaultTooltip();
            chartWpf.DataHover += (sender, chartPoint) =>
            {
                int index = (int)chartPoint.X;
                string numeroObra = labels[index];
                string nomeObra = new MostarGraficos().ObterNomeObra(numeroObra);
                double percentagem = chartPoint.Y;

                labelobraeuros.Text = $"{nomeObra} |  {percentagem:N2} €";
            };
        }
        private void VisualizarGraficoObrasPercentagemAno()
        {
            string anoFecho = ComboBoxAnoAdd.SelectedItem?.ToString();
            if (string.IsNullOrEmpty(anoFecho))
            {
                MessageBox.Show("Selecione um ano para filtrar o gráfico.");
                return;
            }

            var chartWpf = ChartTotalObrasPercentagem.Child as LiveCharts.Wpf.CartesianChart;
            if (chartWpf == null) return;

            var service = new MostarGraficos();
            var (values, labels) = service.CarregarGraficoObrasPercentagemAno(anoFecho);

            chartWpf.Series = new LiveCharts.SeriesCollection
                {
                    new ColumnSeries
                    {
                        Title = "% Total",
                        Values = values,
                        DataLabels = true,
                        Fill = Cor.Brushes.Yellow,
                        Stroke = Cor.Brushes.Black,
                        StrokeThickness = 0.5,
                        LabelPoint = point => point.Y + "%"
                    }
                };

            chartWpf.AxisX.Clear();
            chartWpf.AxisX.Add(new LiveCharts.Wpf.Axis
            {
                Title = "Número da Obra",
                FontSize = 15,
                Labels = labels,
                Separator = new LiveCharts.Wpf.Separator { Step = 1 },
                LabelsRotation = 45
            });

            chartWpf.AxisY.Clear();
            chartWpf.AxisY.Add(new LiveCharts.Wpf.Axis
            {
                Title = "Percentagem (%)",
                LabelFormatter = value => value.ToString("N1") + " %",
                Sections = new SectionsCollection
                     {
                    new AxisSection
                     {
                        Value = 100,
                        Stroke = Cor.Brushes.Red,
                        StrokeThickness = 1,
                        StrokeDashArray = new DoubleCollection { 4 }
                     }
                }
            });
            chartWpf.LegendLocation = LegendLocation.Top;
            chartWpf.Hoverable = true;
            chartWpf.DataTooltip = new DefaultTooltip();
            chartWpf.DataHover += (sender, chartPoint) =>
            {
                int index = (int)chartPoint.X;
                string numeroObra = labels[index];
                string nomeObra = new MostarGraficos().ObterNomeObra(numeroObra);
                double percentagem = chartPoint.Y;

                labelobra.Text = $"{nomeObra} | {percentagem:N1} %";
            };
        }
        private void VisualizarGraficoObrasHorasAno()
        {
            string anoFecho = ComboBoxAnoAdd.SelectedItem?.ToString();
            if (string.IsNullOrEmpty(anoFecho))
            {
                MessageBox.Show("Selecione um ano para filtrar o gráfico.");
                return;
            }

            var chartWpf = chartObrasHoras.Child as LiveCharts.Wpf.CartesianChart;
            if (chartWpf == null) return;

            var service = new MostarGraficos();
            var (orcValues, realValues, labels) = service.CarregarGraficoObrasHorasAno(anoFecho);

            chartWpf.Series = new LiveCharts.SeriesCollection
                    {
                        new ColumnSeries
                        {
                            Title = "Orçamentação",
                            Values = orcValues,
                            DataLabels = true,
                            Fill = Cor.Brushes.LightBlue,
                            Stroke = Cor.Brushes.Black,
                            StrokeThickness = 0.5,
                            LabelPoint = point => point.Y + "h"
                        },
                        new ColumnSeries
                        {
                            Title = "Real",
                            Values = realValues,
                            DataLabels = true,
                            Fill = Cor.Brushes.Orange,
                            Stroke = Cor.Brushes.Black,
                            StrokeThickness = 0.5,
                            LabelPoint = point => point.Y + "h"
                        }
                    };

            chartWpf.AxisX.Clear();
            chartWpf.AxisX.Add(new LiveCharts.Wpf.Axis
            {
                Title = "Número da Obra",
                FontSize = 15,
                Labels = labels,
                Separator = new LiveCharts.Wpf.Separator { Step = 1 },
                LabelsRotation = 45
            });

            chartWpf.AxisY.Clear();
            chartWpf.AxisY.Add(new LiveCharts.Wpf.Axis
            {
                Title = "Horas (h)",
                FontSize = 15,
                LabelFormatter = value => value.ToString("N1") + "h"
            });

            chartWpf.LegendLocation = LegendLocation.Top;
            chartWpf.Hoverable = true;
            chartWpf.DataTooltip = new DefaultTooltip();
            chartWpf.DataHover += (sender, chartPoint) =>
            {
                int index = (int)chartPoint.X;
                string numeroObra = labels[index];
                string nomeObra = new MostarGraficos().ObterNomeObra(numeroObra);
                double percentagem = chartPoint.Y;

                labelobrasHoras.Text = $"{nomeObra} | {percentagem:N1} h";
            };
        }
        private void VisualizarGraficoObrasPercentagemTipologia()
        {
            string tipologia = ComboBoxTipologiaFiltro.SelectedItem?.ToString();
            if (string.IsNullOrEmpty(tipologia))
            {
                MessageBox.Show("Selecione uma Tipologia para filtrar o gráfico.");
                return;
            }

            var chartWpf = ChartTotalObrasPercentagem.Child as LiveCharts.Wpf.CartesianChart;
            if (chartWpf == null) return;

            var service = new MostarGraficos();
            var (values, labels) = service.CarregarGraficoObrasPercentagemTipologia(tipologia);

            chartWpf.Series = new LiveCharts.SeriesCollection
                        {
                            new ColumnSeries
                            {
                                Title = "% Total",
                                Values = values,
                                DataLabels = true,
                                Fill = Cor.Brushes.Yellow,
                                Stroke = Cor.Brushes.Black,
                                StrokeThickness = 0.5,
                                LabelPoint = point => point.Y + "%"
                            }
                        };

            chartWpf.AxisX.Clear();
            chartWpf.AxisX.Add(new LiveCharts.Wpf.Axis
            {
                Title = "Número da Obra",
                FontSize = 15,
                Labels = labels,
                Separator = new LiveCharts.Wpf.Separator { Step = 1 },
                LabelsRotation = 45
            });

            chartWpf.AxisY.Clear();
            chartWpf.AxisY.Add(new LiveCharts.Wpf.Axis
            {
                Title = "Percentagem (%)",
                LabelFormatter = value => value.ToString("N1") + " %",
                Sections = new SectionsCollection
                     {
                    new AxisSection
                     {
                        Value = 100,
                        Stroke = Cor.Brushes.Red,
                        StrokeThickness = 1,
                        StrokeDashArray = new DoubleCollection { 4 }
                     }
                }
            });
            chartWpf.LegendLocation = LegendLocation.Top;
            chartWpf.Hoverable = true;
            chartWpf.DataTooltip = new DefaultTooltip();
            chartWpf.DataHover += (sender, chartPoint) =>
            {
                int index = (int)chartPoint.X;
                string numeroObra = labels[index];
                string nomeObra = new MostarGraficos().ObterNomeObra(numeroObra);
                double percentagem = chartPoint.Y;

                labelobra.Text = $"{nomeObra} | {percentagem:N1} %";
            };
        }
        private void VisualizarGraficoObrasHorasTipologia()
        {
            string tipologia = ComboBoxTipologiaFiltro.SelectedItem?.ToString();
            if (string.IsNullOrEmpty(tipologia))
            {
                MessageBox.Show("Por favor, selecione uma tipologia.");
                return;
            }
            var chartWpf = chartObrasHoras.Child as LiveCharts.Wpf.CartesianChart;
            if (chartWpf == null) return;

            var service = new MostarGraficos();
            var (series, labels) = service.CarregarGraficoObrasHorasTipologia(tipologia);

            chartWpf.Series = series;

            foreach (var s in chartWpf.Series.OfType<ColumnSeries>())
            {
                s.MaxColumnWidth = 30;
                s.ColumnPadding = 10;
            }

            chartWpf.AxisX.Clear();
            chartWpf.AxisX.Add(new LiveCharts.Wpf.Axis
            {
                Title = "Número da Obra",
                FontSize = 15,
                Labels = labels,
                Separator = new LiveCharts.Wpf.Separator { Step = 1 },
                LabelsRotation = 45
            });

            chartWpf.AxisY.Clear();
            chartWpf.AxisY.Add(new LiveCharts.Wpf.Axis
            {
                Title = "Horas (h)",
                FontSize = 15,
                LabelFormatter = value => value + "h"
            });

            chartWpf.LegendLocation = LegendLocation.Top;
            chartWpf.Hoverable = true;
            chartWpf.DataTooltip = new DefaultTooltip();
            chartWpf.DataHover += (sender, chartPoint) =>
            {
                int index = (int)chartPoint.X;
                string numeroObra = labels[index];
                string nomeObra = new MostarGraficos().ObterNomeObra(numeroObra);
                double percentagem = chartPoint.Y;

                labelobrasHoras.Text = $"{nomeObra} | {percentagem:N1} h";
            };
        }
        private void VisualizarGraficoHorasTipologiaeAno()
        {
            var service = new MostarGraficos();
            var (series, labels) = service.CarregarGraficoHorasTotais(
                DataGridViewOrcamentacaoObras, DataGridViewRealObras);

            var chartWpf = chartTotalHoras.Child as LiveCharts.Wpf.CartesianChart;
            if (chartWpf == null) return;

            chartWpf.Series = series;

            chartWpf.AxisX.Clear();
            chartWpf.AxisX.Add(new LiveCharts.Wpf.Axis
            {
                Title = "Total",
                Labels = labels
            });

            chartWpf.AxisY.Clear();
            chartWpf.AxisY.Add(new LiveCharts.Wpf.Axis
            {
                Title = "Horas (h)",
                LabelFormatter = value => value + "h"
            });

            chartWpf.LegendLocation = LegendLocation.Top;
        }
        private void SalvarOrcamentacaoBD()
        {
            if (DataGridViewOrcamentacaoObras.SelectedRows.Count > 0)
            {
                string NumeroObra = DataGridViewOrcamentacaoObras.SelectedRows[0].Cells["Numero da Obra"].Value.ToString();
                string NomeObra = DataGridViewOrcamentacaoObras.SelectedRows[0].Cells["Nome da Obra"].Value.ToString();
                string Preparador = DataGridViewOrcamentacaoObras.SelectedRows[0].Cells["Preparador Responsavel"].Value.ToString();
                string Tipologia = DataGridViewOrcamentacaoObras.SelectedRows[0].Cells["Tipologia"].Value.ToString();
                string KgEstrutura = DataGridViewOrcamentacaoObras.SelectedRows[0].Cells["KG Estrutura"].Value.ToString();
                string HorasEstrutura = DataGridViewOrcamentacaoObras.SelectedRows[0].Cells["Horas Estrutura"].Value.ToString();
                string HorasRevestimentos = DataGridViewOrcamentacaoObras.SelectedRows[0].Cells["Horas Revestimentos"].Value.ToString();

                string LabelPrecoOrcamen = LabelPrecoOrcamentado.Text;
                double precoOrcamentadoDouble = Convert.ToDouble(LabelPrecoOrcamen);

                double horasEstruturaDouble = Convert.ToDouble(HorasEstrutura);
                string horasEstruturaString = horasEstruturaDouble.ToString("F1");
                double horasEstruturaDouble2 = Convert.ToDouble(horasEstruturaString);
                double ValorEstrutura = horasEstruturaDouble2 * precoOrcamentadoDouble;

                string valorEstruturaString = ValorEstrutura.ToString("F2");
                double valorEstruturaDouble = Convert.ToDouble(valorEstruturaString);

                double kgEstruturaDouble = Convert.ToDouble(KgEstrutura);
                double kgEuroEstrutura = valorEstruturaDouble / kgEstruturaDouble;
                double kgEuroEstruturaArredondado = Math.Round(kgEuroEstrutura, 2);
                string KgEuroEstruturaString = kgEuroEstruturaArredondado.ToString("F2");

                double HorasRevestimentosDouble = Convert.ToDouble(HorasRevestimentos);
                double ValorRevestimentos = HorasRevestimentosDouble * precoOrcamentadoDouble;
                double ValorRevestimentosArredondado = Math.Round(ValorRevestimentos, 2);
                string valorRevestimentosString = ValorRevestimentosArredondado.ToString("F2");
                double TotalHoras = horasEstruturaDouble + HorasRevestimentosDouble;

                double TotalValor = valorEstruturaDouble + ValorRevestimentos;
                double TotalValorArredondado = Math.Round(TotalValor, 2);
                string TotalValorString = TotalValorArredondado.ToString("F2");


                string queryOrcamentacao = "INSERT INTO dbo.Orçamentação ([Numero da Obra], [Nome da Obra], [Preparador Responsavel], Tipologia, [KG Estrutura], [Horas Estrutura], [Valor Estrutura], [KG/Euro Estrutura], [Horas Revestimentos], [Valor Revestimentos], [Total Horas], [Total Valor]) " +
                                           "VALUES (@NumeroObra, @NomeObra, @PreparadorResponsavel, @Tipologia, @KGEstrutura, @HorasEstrutura, @ValorEstrutura, @KGEuroEstrutura, @HorasRevestimentos, @ValorRevestimentos, @TotalHoras, @TotalValor)";

                string queryRealObras = "INSERT INTO dbo.RealObras ([Numero da Obra], [Nome da Obra], [Preparador Responsavel], Tipologia, [KG Estrutura], [Horas Estrutura], [Valor Estrutura], [KG/Euro Estrutura], [Percentagem Estrutura], [Horas Revestimentos], [Valor Revestimentos], [Percentagem Revestimentos], [Horas Aprovação], [Valor Aprovação], [Percentagem Aprovação], [Horas Alterações], [Valor Alterações], [Percentagem Alterações], [Horas Fabrico], [Valor Fabrico], [Percentagem Fabrico], [Horas Soldadura], [Valor Soldadura], [Percentagem Soldadura], [Horas Montagem], [Valor Montagem], [Percentagem Montagem], [Horas Diversos], [Valor Diversos], [Percentagem Diversos], [Comentario Diversos], [Total Horas], [Total Valor]) " +
                                        "VALUES (@NumeroObra, @NomeObra, @PreparadorResponsavel, @Tipologia, @KGEstrutura, 0, 0, @KGEuroEstrutura, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)";

                string queryConclusaoObras = "INSERT INTO dbo.ConclusaoObras ([Numero da Obra], [Nome da Obra], [Preparador Responsavel], Tipologia, [Total Horas], [Total Valor], [Percentagem Total], [Dias de Preparação]) " +
                                             "VALUES (@NumeroObra, @NomeObra, @PreparadorResponsavel, @Tipologia, 0, 0, 0, 0)";

                ComunicaBD BD = new ComunicaBD();
                try
                {
                    BD.ConectarBD();
                    using (SqlCommand cmd = new SqlCommand(queryOrcamentacao, BD.GetConnection()))
                    {
                        cmd.Parameters.AddWithValue("@NumeroObra", NumeroObra);
                        cmd.Parameters.AddWithValue("@NomeObra", NomeObra);
                        cmd.Parameters.AddWithValue("@PreparadorResponsavel", Preparador);
                        cmd.Parameters.AddWithValue("@Tipologia", Tipologia);
                        cmd.Parameters.AddWithValue("@KGEstrutura", KgEstrutura);
                        cmd.Parameters.AddWithValue("@HorasEstrutura", horasEstruturaString);
                        cmd.Parameters.AddWithValue("@ValorEstrutura", valorEstruturaString);
                        cmd.Parameters.AddWithValue("@KGEuroEstrutura", KgEuroEstruturaString);
                        cmd.Parameters.AddWithValue("@HorasRevestimentos", HorasRevestimentos);
                        cmd.Parameters.AddWithValue("@ValorRevestimentos", valorRevestimentosString);
                        cmd.Parameters.AddWithValue("@TotalHoras", TotalHoras);
                        cmd.Parameters.AddWithValue("@TotalValor", TotalValorString);
                        cmd.ExecuteNonQuery();
                    }
                    using (SqlCommand cmdRealObras = new SqlCommand(queryRealObras, BD.GetConnection()))
                    {
                        cmdRealObras.Parameters.AddWithValue("@NumeroObra", NumeroObra);
                        cmdRealObras.Parameters.AddWithValue("@NomeObra", NomeObra);
                        cmdRealObras.Parameters.AddWithValue("@PreparadorResponsavel", Preparador);
                        cmdRealObras.Parameters.AddWithValue("@Tipologia", Tipologia);
                        cmdRealObras.Parameters.AddWithValue("@KGEstrutura", KgEstrutura);
                        cmdRealObras.Parameters.AddWithValue("@KGEuroEstrutura", KgEuroEstruturaString);
                        cmdRealObras.ExecuteNonQuery();
                    }
                    using (SqlCommand cmdConclusaoObras = new SqlCommand(queryConclusaoObras, BD.GetConnection()))
                    {
                        cmdConclusaoObras.Parameters.AddWithValue("@NumeroObra", NumeroObra);
                        cmdConclusaoObras.Parameters.AddWithValue("@NomeObra", NomeObra);
                        cmdConclusaoObras.Parameters.AddWithValue("@PreparadorResponsavel", Preparador);
                        cmdConclusaoObras.Parameters.AddWithValue("@Tipologia", Tipologia);
                        cmdConclusaoObras.ExecuteNonQuery();
                    }
                    MessageBox.Show("Valores registrados com sucesso em todas as tabelas ( Real / Conclusão ).");
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
        private void AtualizarOrcamentoNaBD()
        {
            if (DataGridViewOrcamentacaoObras.SelectedRows.Count > 0)
            {
                string ID = DataGridViewOrcamentacaoObras.SelectedRows[0].Cells["ID"].Value.ToString();
                string NumeroObra = DataGridViewOrcamentacaoObras.SelectedRows[0].Cells["Numero da Obra"].Value.ToString();
                string NomeObra = DataGridViewOrcamentacaoObras.SelectedRows[0].Cells["Nome da Obra"].Value.ToString();
                string Preparador = DataGridViewOrcamentacaoObras.SelectedRows[0].Cells["Preparador Responsavel"].Value.ToString();
                string Tipologia = DataGridViewOrcamentacaoObras.SelectedRows[0].Cells["Tipologia"].Value.ToString();
                string KgEstrutura = DataGridViewOrcamentacaoObras.SelectedRows[0].Cells["KG Estrutura"].Value.ToString();
                string HorasEstrutura = DataGridViewOrcamentacaoObras.SelectedRows[0].Cells["Horas Estrutura"].Value.ToString();
                string HorasRevestimentos = DataGridViewOrcamentacaoObras.SelectedRows[0].Cells["Horas Revestimentos"].Value.ToString();

                string LabelPrecoOrcamen = LabelPrecoOrcamentado.Text;
                double precoOrcamentadoDouble = Convert.ToDouble(LabelPrecoOrcamen);
                double horasEstruturaDouble = Convert.ToDouble(HorasEstrutura);
                double ValorEstrutura = horasEstruturaDouble * precoOrcamentadoDouble;
                string valorEstruturaString = ValorEstrutura.ToString("F2");
                double valorEstruturaDouble = Convert.ToDouble(valorEstruturaString);
                double kgEstruturaDouble = Convert.ToDouble(KgEstrutura);
                double kgEuroEstrutura = valorEstruturaDouble / kgEstruturaDouble;
                double kgEuroEstruturaArredondado = Math.Round(kgEuroEstrutura, 2);
                string KgEuroEstruturaString = kgEuroEstruturaArredondado.ToString("F2");
                double HorasRevestimentosDouble = Convert.ToDouble(HorasRevestimentos);
                double ValorRevestimentos = HorasRevestimentosDouble * precoOrcamentadoDouble;
                double ValorRevestimentosArredondado = Math.Round(ValorRevestimentos, 2);
                string valorRevestimentosString = ValorRevestimentosArredondado.ToString("F2");
                double TotalHoras = horasEstruturaDouble + HorasRevestimentosDouble;
                double TotalValor = valorEstruturaDouble + ValorRevestimentos;
                double TotalValorArredondado = Math.Round(TotalValor, 2);
                string TotalValorString = TotalValorArredondado.ToString("F2");

                string query = "UPDATE dbo.Orçamentação " +
                               "SET [Numero da Obra] = @NumeroObra, " +
                               "[Nome da Obra] = @NomeObra, " +
                               "[Preparador Responsavel] = @PreparadorResponsavel, " +
                               "Tipologia = @Tipologia, " +
                               "[KG Estrutura] = @KGEstrutura, " +
                               "[Horas Estrutura] = @HorasEstrutura, " +
                               "[Valor Estrutura] = @ValorEstrutura, " +
                               "[KG/Euro Estrutura] = @KGEuroEstrutura, " +
                               "[Horas Revestimentos] = @HorasRevestimentos, " +
                               "[Valor Revestimentos] = @ValorRevestimentos, " +
                               "[Total Horas] = @TotalHoras, " +
                               "[Total Valor] = @TotalValor " +
                               "WHERE ID = @ID";

                ComunicaBD BD = new ComunicaBD();
                try
                {
                    BD.ConectarBD();

                    using (SqlCommand cmd = new SqlCommand(query, BD.GetConnection()))
                    {
                        cmd.Parameters.AddWithValue("@NumeroObra", NumeroObra);
                        cmd.Parameters.AddWithValue("@NomeObra", NomeObra);
                        cmd.Parameters.AddWithValue("@PreparadorResponsavel", Preparador);
                        cmd.Parameters.AddWithValue("@Tipologia", Tipologia);
                        cmd.Parameters.AddWithValue("@KGEstrutura", KgEstrutura);
                        cmd.Parameters.AddWithValue("@HorasEstrutura", HorasEstrutura);
                        cmd.Parameters.AddWithValue("@ValorEstrutura", valorEstruturaString);
                        cmd.Parameters.AddWithValue("@KGEuroEstrutura", KgEuroEstruturaString);
                        cmd.Parameters.AddWithValue("@HorasRevestimentos", HorasRevestimentos);
                        cmd.Parameters.AddWithValue("@ValorRevestimentos", valorRevestimentosString);
                        cmd.Parameters.AddWithValue("@TotalHoras", TotalHoras);
                        cmd.Parameters.AddWithValue("@TotalValor", TotalValorString);
                        cmd.Parameters.AddWithValue("@ID", ID);

                        cmd.ExecuteNonQuery();
                    }

                    MessageBox.Show("Valores Atualizados com Sucesso.");
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
        private void AtualizarTabelaRealnaBd()
        {
            if (DataGridViewRealObras.Rows.Count > 0)
            {
                ComunicaBD BD = new ComunicaBD();
                try
                {
                    BD.ConectarBD();

                    foreach (DataGridViewRow row in DataGridViewRealObras.Rows)
                    {
                        if (row.Cells["Numero da Obra"].Value != null)
                        {
                            string NumeroObra = row.Cells["Numero da Obra"].Value.ToString();
                            string ID = row.Cells["ID"].Value.ToString();

                            double horasEstruturaTotal1 = 0;
                            double horasRevestimentosTotal1 = 0;
                            double horasAprovacaoTotal1 = 0;
                            double horasAlteracoesTotal1 = 0;
                            double horasFabricoTotal1 = 0;
                            double horasSoldaduraTotal1 = 0;
                            double horasMontagemTotal1 = 0;
                            double horasDiversasTotal1 = 0;

                            horasEstruturaTotal1 = ObterHoras(NumeroObra, 401, BD);
                            horasRevestimentosTotal1 = ObterHoras(NumeroObra, 407, BD);
                            horasAprovacaoTotal1 = ObterHoras(NumeroObra, 402, BD);
                            horasAlteracoesTotal1 = ObterHoras(NumeroObra, 409, BD);
                            horasFabricoTotal1 = ObterHoras(NumeroObra, 403, BD);
                            horasSoldaduraTotal1 = ObterHoras(NumeroObra, 404, BD);
                            horasMontagemTotal1 = ObterHoras(NumeroObra, 413, BD);
                            horasDiversasTotal1 = ObterHoras(NumeroObra, 408, BD);

                            int horasEstruturaTotal = Convert.ToInt32(Math.Floor(horasEstruturaTotal1 / 60));
                            int horasRevestimentosTotal = Convert.ToInt32(Math.Floor(horasRevestimentosTotal1 / 60));
                            int horasAprovacaoTotal = Convert.ToInt32(Math.Floor(horasAprovacaoTotal1 / 60));
                            int horasAlteracoesTotal = Convert.ToInt32(Math.Floor(horasAlteracoesTotal1 / 60));
                            int horasFabricoTotal = Convert.ToInt32(Math.Floor(horasFabricoTotal1 / 60));
                            int horasSoldaduraTotal = Convert.ToInt32(Math.Floor(horasSoldaduraTotal1 / 60));
                            int horasMontagemTotal = Convert.ToInt32(Math.Floor(horasMontagemTotal1 / 60));
                            int horasDiversasTotal = Convert.ToInt32(Math.Floor(horasDiversasTotal1 / 60));

                            int TotalHoras = horasEstruturaTotal + horasRevestimentosTotal + horasAprovacaoTotal +
                                             horasAlteracoesTotal + horasFabricoTotal + horasSoldaduraTotal + horasMontagemTotal + horasDiversasTotal;


                            string updateQuery = "UPDATE dbo.RealObras " +
                                                    "SET [Numero da Obra] = @NumeroObra, " +
                                                    "[Horas Estrutura] = @HorasEstrutura, " +
                                                    "[Horas Revestimentos] = @HorasRevestimentos, " +
                                                    "[Horas Aprovação] = @HorasAprovacao, " +
                                                    "[Horas Alterações] = @HorasAlteracoes, " +
                                                    "[Horas Fabrico] = @HorasFabrico, " +
                                                    "[Horas Soldadura] = @HorasSoldadura, " +
                                                    "[Horas Montagem] = @HorasMontagem, " +
                                                    "[Horas Diversos] = @HorasDiversos, " +
                                                    "[Total Horas] = @TotalHoras " +
                                                    "WHERE ID = @ID";

                            using (SqlCommand updateCmd = new SqlCommand(updateQuery, BD.GetConnection()))
                            {
                                updateCmd.Parameters.AddWithValue("@NumeroObra", NumeroObra);
                                updateCmd.Parameters.AddWithValue("@HorasEstrutura", horasEstruturaTotal);
                                updateCmd.Parameters.AddWithValue("@HorasRevestimentos", horasRevestimentosTotal);
                                updateCmd.Parameters.AddWithValue("@HorasAprovacao", horasAprovacaoTotal);
                                updateCmd.Parameters.AddWithValue("@HorasAlteracoes", horasAlteracoesTotal);
                                updateCmd.Parameters.AddWithValue("@HorasFabrico", horasFabricoTotal);
                                updateCmd.Parameters.AddWithValue("@HorasSoldadura", horasSoldaduraTotal);
                                updateCmd.Parameters.AddWithValue("@HorasMontagem", horasMontagemTotal);
                                updateCmd.Parameters.AddWithValue("@HorasDiversos", horasDiversasTotal);
                                updateCmd.Parameters.AddWithValue("@TotalHoras", TotalHoras);
                                updateCmd.Parameters.AddWithValue("@ID", ID);

                                int rowsAffected = updateCmd.ExecuteNonQuery();
                            }
                        }
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
        private double ObterHoras(string numeroObra, int codigoTarefa, ComunicaBD BD)
        {
            string query = @"
                            SELECT 
                                SUM(DATEDIFF(MINUTE, '00:00:00', TRY_CAST([Qtd de Hora] AS TIME))) AS TotalHorasEmSegundos
                            FROM 
                                dbo.RegistoTempo
                            WHERE 
                                [Numero da Obra] = @NumeroObra
                                AND [Codigo da Tarefa] = @CodigoTarefa";

            using (SqlCommand cmd = new SqlCommand(query, BD.GetConnection()))
            {
                cmd.Parameters.AddWithValue("@NumeroObra", numeroObra);
                cmd.Parameters.AddWithValue("@CodigoTarefa", codigoTarefa);
                return cmd.ExecuteScalar() == DBNull.Value ? 0 : Convert.ToDouble(cmd.ExecuteScalar());
            }
        }
        private void AtualizarTabelaRealnaBd2()
        {
            if (DataGridViewRealObras.Rows.Count > 0)
            {
                ComunicaBD BD = new ComunicaBD();
                try
                {
                    BD.ConectarBD();
                    foreach (DataGridViewRow row in DataGridViewRealObras.Rows)
                    {
                        if (row.IsNewRow) continue;
                        string ID = row.Cells["ID"].Value.ToString();
                        string NumeroObra = row.Cells["Numero da Obra"].Value.ToString();
                        string HorasEstrutura = row.Cells["Horas Estrutura"].Value.ToString().Replace(",", ".").Trim();
                        string HorasRevestimentos = row.Cells["Horas Revestimentos"].Value.ToString().Replace(",", ".").Trim();
                        string HorasAprovacao = row.Cells["Horas Aprovação"].Value.ToString().Replace(",", ".").Trim();
                        string HorasAlteracoes = row.Cells["Horas Alterações"].Value.ToString().Replace(",", ".").Trim();
                        string HorasFabrico = row.Cells["Horas Fabrico"].Value.ToString().Replace(",", ".").Trim();
                        string HorasSoldadura = row.Cells["Horas Soldadura"].Value.ToString().Replace(",", ".").Trim();
                        string HorasMontagem = row.Cells["Horas Montagem"].Value.ToString().Replace(",", ".").Trim();
                        string HorasDiversos = row.Cells["Horas Diversos"].Value.ToString().Replace(",", ".").Trim();

                        double PrecoOrcamentado = 0;
                        if (!double.TryParse(LabelPrecoOrcamentado.Text, out PrecoOrcamentado))
                        {
                            MessageBox.Show("Preço Orçamentado inválido.");
                            return;
                        }

                        CultureInfo culture = CultureInfo.InvariantCulture;

                        double ValorEstrutura = 0;
                        if (double.TryParse(HorasEstrutura, NumberStyles.Any, culture, out double HorasEstruturaT))
                        {
                            ValorEstrutura = HorasEstruturaT * PrecoOrcamentado;
                        }
                        else
                        {
                            MessageBox.Show($"Erro ao converter Horas Estrutura: '{HorasEstrutura}'");
                        }

                        double ValorRevestimentos = 0;
                        if (double.TryParse(HorasRevestimentos, NumberStyles.Any, culture, out double HorasRevestimentosT))
                        {
                            ValorRevestimentos = HorasRevestimentosT * PrecoOrcamentado;
                        }
                        else
                        {
                            MessageBox.Show($"Erro ao converter Horas Revestimentos: '{HorasRevestimentos}'");
                        }

                        double ValorAprovacao = 0;
                        if (double.TryParse(HorasAprovacao, NumberStyles.Any, culture, out double HorasAprovacaoT))
                        {
                            ValorAprovacao = HorasAprovacaoT * PrecoOrcamentado;
                        }
                        else
                        {
                            MessageBox.Show($"Erro ao converter Horas Aprovação: '{HorasAprovacao}'");
                        }

                        double ValorAlteracoes = 0;
                        if (double.TryParse(HorasAlteracoes, NumberStyles.Any, culture, out double HorasAlteracoesT))
                        {
                            ValorAlteracoes = HorasAlteracoesT * PrecoOrcamentado;
                        }
                        else
                        {
                            MessageBox.Show($"Erro ao converter Horas Alterações: '{HorasAlteracoes}'");
                        }

                        double ValorFabrico = 0;
                        if (double.TryParse(HorasFabrico, NumberStyles.Any, culture, out double HorasFabricoT))
                        {
                            ValorFabrico = HorasFabricoT * PrecoOrcamentado;
                        }
                        else
                        {
                            MessageBox.Show($"Erro ao converter Horas Fabrico: '{HorasFabrico}'");
                        }

                        double ValorSoldadura = 0;
                        if (double.TryParse(HorasSoldadura, NumberStyles.Any, culture, out double HorasSoldaduraT))
                        {
                            ValorSoldadura = HorasSoldaduraT * PrecoOrcamentado;
                        }
                        else
                        {
                            MessageBox.Show($"Erro ao converter Horas Soldadura: '{HorasSoldadura}'");
                        }

                        double ValorMontagem = 0;
                        if (double.TryParse(HorasMontagem, NumberStyles.Any, culture, out double HorasMontagemT))
                        {
                            ValorMontagem = HorasMontagemT * PrecoOrcamentado;
                        }
                        else
                        {
                            MessageBox.Show($"Erro ao converter Horas Montagem: '{HorasMontagem}'");
                        }

                        double ValorDiversos = 0;
                        if (double.TryParse(HorasDiversos, NumberStyles.Any, culture, out double HorasDiversosT))
                        {
                            ValorDiversos = HorasDiversosT * PrecoOrcamentado;
                        }
                        else
                        {
                            MessageBox.Show($"Erro ao converter Horas Diversos: '{HorasDiversos}'");
                        }

                        double ValorTotalHoras = ValorEstrutura + ValorRevestimentos + ValorAprovacao + ValorAlteracoes + ValorFabrico + ValorSoldadura + ValorMontagem + ValorDiversos;

                        string query = "UPDATE dbo.RealObras " +
                                        "SET [Numero da Obra] = @NumeroObra, " +
                                        "[Valor Estrutura] = @ValorEstrutura, " +
                                        "[Valor Revestimentos] = @ValorRevestimentos, " +
                                        "[Valor Aprovação] = @ValorAprovacao, " +
                                        "[Valor Alterações] = @ValorAlteracoes, " +
                                        "[Valor Fabrico] = @ValorFabrico, " +
                                        "[Valor Soldadura] = @ValorSoldadura, " +
                                        "[Valor Montagem] = @ValorMontagem, " +
                                        "[Valor Diversos] = @ValorDiversos, " +
                                        "[Total Valor] = @TotalValor " +
                                        "WHERE ID = @ID";

                        using (SqlCommand cmd = new SqlCommand(query, BD.GetConnection()))
                        {
                            cmd.Parameters.AddWithValue("@NumeroObra", NumeroObra);
                            cmd.Parameters.AddWithValue("@ID", ID);
                            cmd.Parameters.AddWithValue("@ValorEstrutura", ValorEstrutura);
                            cmd.Parameters.AddWithValue("@ValorRevestimentos", ValorRevestimentos);
                            cmd.Parameters.AddWithValue("@ValorAprovacao", ValorAprovacao);
                            cmd.Parameters.AddWithValue("@ValorAlteracoes", ValorAlteracoes);
                            cmd.Parameters.AddWithValue("@ValorFabrico", ValorFabrico);
                            cmd.Parameters.AddWithValue("@ValorSoldadura", ValorSoldadura);
                            cmd.Parameters.AddWithValue("@ValorMontagem", ValorMontagem);
                            cmd.Parameters.AddWithValue("@ValorDiversos", ValorDiversos);
                            cmd.Parameters.AddWithValue("@TotalValor", ValorTotalHoras);

                            cmd.ExecuteNonQuery();
                        }
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
            else
            {
                MessageBox.Show("Não há dados na tabela.");
            }
        }
        private void AtualizarTabelaRealnaBd3()
        {
            if (DataGridViewRealObras.Rows.Count > 0)
            {
                ComunicaBD BD = new ComunicaBD();
                try
                {
                    BD.ConectarBD();

                    foreach (DataGridViewRow row in DataGridViewRealObras.Rows)
                    {
                        if (row.IsNewRow) continue;

                        string ID = row.Cells["ID"].Value.ToString();
                        string NumeroObra = row.Cells["Numero da Obra"].Value.ToString();

                        string HorasEstrutura = row.Cells["Horas Estrutura"].Value.ToString();
                        string HorasRevestimentos = row.Cells["Horas Revestimentos"].Value.ToString();
                        string HorasAprovacao = row.Cells["Horas Aprovação"].Value.ToString();
                        string HorasAlteracoes = row.Cells["Horas Alterações"].Value.ToString();
                        string HorasFabrico = row.Cells["Horas Fabrico"].Value.ToString();
                        string HorasSoldadura = row.Cells["Horas Soldadura"].Value.ToString();
                        string HorasMontagem = row.Cells["Horas Montagem"].Value.ToString();
                        string HorasDiversos = row.Cells["Horas Diversos"].Value.ToString();
                        string HorasHoras = row.Cells["Total Horas"].Value.ToString();

                        double HorasHorasP = 0;
                        double HorasEstruturaP = 0;
                        double HorasRevestimentosP = 0;
                        double HorasAprovacaoP = 0;
                        double HorasAlteracoesP = 0;
                        double HorasFabricoP = 0;
                        double HorasSoldaduraP = 0;
                        double HorasMontagemP = 0;
                        double HorasDiversosP = 0;

                        double PercentagemEstrutura = 0;
                        double PercentagemRevestimentos = 0;
                        double PercentagemAprovacao = 0;
                        double PercentagemAlteracoes = 0;
                        double PercentagemFabrico = 0;
                        double PercentagemSoldadura = 0;
                        double PercentagemMontagem = 0;
                        double PercentagemDiversos = 0;

                        HorasHorasP = double.Parse(HorasHoras, CultureInfo.InvariantCulture);
                        HorasEstruturaP = double.Parse(HorasEstrutura, CultureInfo.InvariantCulture);
                        HorasRevestimentosP = double.Parse(HorasRevestimentos, CultureInfo.InvariantCulture);
                        HorasAprovacaoP = double.Parse(HorasAprovacao, CultureInfo.InvariantCulture);
                        HorasAlteracoesP = double.Parse(HorasAlteracoes, CultureInfo.InvariantCulture);
                        HorasFabricoP = double.Parse(HorasFabrico, CultureInfo.InvariantCulture);
                        HorasSoldaduraP = double.Parse(HorasSoldadura, CultureInfo.InvariantCulture);
                        HorasMontagemP = double.Parse(HorasMontagem, CultureInfo.InvariantCulture);
                        HorasDiversosP = double.Parse(HorasDiversos, CultureInfo.InvariantCulture);

                        if (HorasHorasP != 0)
                        {
                            PercentagemEstrutura = (HorasEstruturaP / HorasHorasP) * 100;
                        }
                        if (HorasHorasP != 0)
                        {
                            PercentagemRevestimentos = (HorasRevestimentosP / HorasHorasP) * 100;
                        }
                        if (HorasHorasP != 0)
                        {
                            PercentagemAprovacao = (HorasAprovacaoP / HorasHorasP) * 100;
                        }
                        if (HorasHorasP != 0)
                        {
                            PercentagemAlteracoes = (HorasAlteracoesP / HorasHorasP) * 100;
                        }
                        if (HorasHorasP != 0)
                        {
                            PercentagemFabrico = (HorasFabricoP / HorasHorasP) * 100;
                        }
                        if (HorasHorasP != 0)
                        {
                            PercentagemSoldadura = (HorasSoldaduraP / HorasHorasP) * 100;
                        }
                        if (HorasHorasP != 0)
                        {
                            PercentagemMontagem = (HorasMontagemP / HorasHorasP) * 100;
                        }
                        if (HorasHorasP != 0)
                        {
                            PercentagemDiversos = (HorasDiversosP / HorasHorasP) * 100;
                        }

                        PercentagemEstrutura = Math.Round(PercentagemEstrutura, 1);
                        PercentagemRevestimentos = Math.Round(PercentagemRevestimentos, 1);
                        PercentagemAprovacao = Math.Round(PercentagemAprovacao, 1);
                        PercentagemAlteracoes = Math.Round(PercentagemAlteracoes, 1);
                        PercentagemFabrico = Math.Round(PercentagemFabrico, 1);
                        PercentagemSoldadura = Math.Round(PercentagemSoldadura, 1);
                        PercentagemMontagem = Math.Round(PercentagemMontagem, 1);
                        PercentagemDiversos = Math.Round(PercentagemDiversos, 1);

                        string PercentagemEstruturaStr = PercentagemEstrutura.ToString("0.0") + "%";
                        string PercentagemRevestimentosStr = PercentagemRevestimentos.ToString("0.0") + "%";
                        string PercentagemAprovacaoStr = PercentagemAprovacao.ToString("0.0") + "%";
                        string PercentagemAlteracoesStr = PercentagemAlteracoes.ToString("0.0") + "%";
                        string PercentagemFabricoStr = PercentagemFabrico.ToString("0.0") + "%";
                        string PercentagemSoldaduraStr = PercentagemSoldadura.ToString("0.0") + "%";
                        string PercentagemMontagemStr = PercentagemMontagem.ToString("0.0") + "%";
                        string PercentagemDiversosStr = PercentagemDiversos.ToString("0.0") + "%";

                        string query = "UPDATE dbo.RealObras " +
                                       "SET [Numero da Obra] = @NumeroObra, " +
                                       "[Percentagem Estrutura] = @PercentagemEstrutura, " +
                                       "[Percentagem Revestimentos] = @PercentagemRevestimentos, " +
                                       "[Percentagem Aprovação] = @PercentagemAprovacao, " +
                                       "[Percentagem Alterações] = @PercentagemAlteracoes, " +
                                       "[Percentagem Fabrico] = @PercentagemFabrico, " +
                                       "[Percentagem Soldadura] = @PercentagemSoldadura, " +
                                       "[Percentagem Montagem] = @PercentagemMontagem, " +
                                       "[Percentagem Diversos] = @PercentagemDiversos " +
                                       "WHERE ID = @ID";

                        using (SqlCommand cmd = new SqlCommand(query, BD.GetConnection()))
                        {
                            cmd.Parameters.AddWithValue("@NumeroObra", NumeroObra);
                            cmd.Parameters.AddWithValue("@ID", ID);
                            cmd.Parameters.AddWithValue("@PercentagemEstrutura", PercentagemEstruturaStr);
                            cmd.Parameters.AddWithValue("@PercentagemRevestimentos", PercentagemRevestimentosStr);
                            cmd.Parameters.AddWithValue("@PercentagemAprovacao", PercentagemAprovacaoStr);
                            cmd.Parameters.AddWithValue("@PercentagemAlteracoes", PercentagemAlteracoesStr);
                            cmd.Parameters.AddWithValue("@PercentagemFabrico", PercentagemFabricoStr);
                            cmd.Parameters.AddWithValue("@PercentagemSoldadura", PercentagemSoldaduraStr);
                            cmd.Parameters.AddWithValue("@PercentagemMontagem", PercentagemMontagemStr);
                            cmd.Parameters.AddWithValue("@PercentagemDiversos", PercentagemDiversosStr);
                            cmd.ExecuteNonQuery();
                        }
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
            else
            {
                MessageBox.Show("Não há dados na tabela.");
            }
        }
        private void AtualizarTabelaConclusaoBD()
        {
            if (DataGridViewConclusaoObras.Rows.Count > 0)
            {
                ComunicaBD BD = new ComunicaBD();
                try
                {
                    BD.ConectarBD();

                    foreach (DataGridViewRow row in DataGridViewConclusaoObras.Rows)
                    {
                        if (row.Cells["Numero da Obra"].Value != null && !string.IsNullOrWhiteSpace(row.Cells["Numero da Obra"].Value.ToString()))
                        {
                            string NumeroObra = row.Cells["Numero da Obra"].Value.ToString();
                            string ID = row.Cells["ID"].Value.ToString();
                            double TotalHorasReal = 0;
                            double TotalHorasOrcamentacao = 0;
                            double TotalHorasResultado = 0;
                            double PercentagemTotal = 0;
                            double DiasPreparar = 0;
                            string queryRealObras = "SELECT [Total Horas] FROM dbo.RealObras WHERE [Numero da Obra] = @NumeroObra";
                            using (SqlCommand cmd = new SqlCommand(queryRealObras, BD.GetConnection()))
                            {
                                cmd.Parameters.AddWithValue("@NumeroObra", NumeroObra);

                                using (SqlDataReader reader = cmd.ExecuteReader())
                                {
                                    if (reader.Read())
                                    {
                                        if (reader["Total Horas"] != DBNull.Value)
                                        {
                                            string valorRealHoras = reader["Total Horas"].ToString();

                                            if (double.TryParse(valorRealHoras, NumberStyles.Any, CultureInfo.InvariantCulture, out TotalHorasReal))
                                            { }
                                            else
                                            {
                                                MessageBox.Show($"Erro ao converter o valor de 'Total Horas' para Número da Obra {NumeroObra}. Valor: {valorRealHoras}");
                                            }
                                        }
                                        else
                                        {
                                            MessageBox.Show($"Valor de 'Total Horas' não encontrado ou é nulo para o Número da Obra {NumeroObra}.");
                                        }
                                    }
                                    else
                                    {
                                        MessageBox.Show($"Número da obra {NumeroObra} não encontrado na tabela RealObras.");
                                    }
                                }
                            }
                            string queryOrcamentacao = "SELECT [Total Horas] FROM dbo.Orçamentação WHERE [Numero da Obra] = @NumeroObra";
                            using (SqlCommand cmd = new SqlCommand(queryOrcamentacao, BD.GetConnection()))
                            {
                                cmd.Parameters.AddWithValue("@NumeroObra", NumeroObra);

                                using (SqlDataReader reader = cmd.ExecuteReader())
                                {
                                    if (reader.Read())
                                    {
                                        if (reader["Total Horas"] != DBNull.Value)
                                        {
                                            string valorOrcamentoHoras = reader["Total Horas"].ToString();

                                            if (double.TryParse(valorOrcamentoHoras, NumberStyles.Any, CultureInfo.InvariantCulture, out TotalHorasOrcamentacao))
                                            { }
                                            else
                                            {
                                                MessageBox.Show($"Erro ao converter o valor de 'Total Horas' para Número da Obra {NumeroObra}. Valor: {valorOrcamentoHoras}");
                                            }
                                        }
                                        else
                                        {
                                            MessageBox.Show($"Valor de 'Total Horas' não encontrado ou é nulo para o Número da Obra {NumeroObra}.");
                                        }
                                    }
                                    else
                                    {
                                        MessageBox.Show($"Número da obra {NumeroObra} não encontrado na tabela Orçamentação.");
                                    }
                                }
                            }
                            TotalHorasResultado = TotalHorasReal - TotalHorasOrcamentacao;
                            PercentagemTotal = (TotalHorasReal / TotalHorasOrcamentacao) * 100;
                            PercentagemTotal = Math.Round(PercentagemTotal, 1);
                            string PercentagemT = PercentagemTotal.ToString("0.0") + "%";
                            DiasPreparar = TotalHorasResultado / 8;
                            int DiasPrepararInteiro = (int)Math.Round(DiasPreparar);

                            string updateQuery = "UPDATE dbo.ConclusaoObras " +
                                                 "SET [Total Horas] = @TotalHoras, " +
                                                 "[Percentagem Total] = @PercentagemT, " +
                                                 "[Dias de Preparação] = @DiasP " +
                                                 "WHERE [Numero da Obra] = @NumeroObra";

                            using (SqlCommand updateCmd = new SqlCommand(updateQuery, BD.GetConnection()))
                            {
                                updateCmd.Parameters.AddWithValue("@TotalHoras", TotalHorasResultado);
                                updateCmd.Parameters.AddWithValue("@PercentagemT", PercentagemT);
                                updateCmd.Parameters.AddWithValue("@DiasP", DiasPrepararInteiro);
                                updateCmd.Parameters.AddWithValue("@NumeroObra", NumeroObra);

                                int rowsAffected = updateCmd.ExecuteNonQuery();

                            }
                        }
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
            else
            {
                MessageBox.Show("Não há dados na tabela.");
            }
        }
        private void AtualizarTabelaConclusaoBD2()
        {
            if (DataGridViewConclusaoObras.Rows.Count > 0)
            {
                ComunicaBD BD = new ComunicaBD();
                try
                {
                    BD.ConectarBD();

                    foreach (DataGridViewRow row in DataGridViewConclusaoObras.Rows)
                    {
                        if (row.Cells["Numero da Obra"].Value != null && !string.IsNullOrWhiteSpace(row.Cells["Numero da Obra"].Value.ToString()))
                        {
                            string NumeroObra = row.Cells["Numero da Obra"].Value.ToString();
                            string ID = row.Cells["ID"].Value.ToString();
                            double TotalValorReal = 0;
                            double TotalValorOrcamentacao = 0;
                            double TotalValorResultado = 0;
                            string queryRealObrasValor = "SELECT [Total Valor] FROM dbo.RealObras WHERE [Numero da Obra] = @NumeroObra";
                            using (SqlCommand cmd = new SqlCommand(queryRealObrasValor, BD.GetConnection()))
                            {
                                cmd.Parameters.AddWithValue("@NumeroObra", NumeroObra);

                                using (SqlDataReader reader = cmd.ExecuteReader())
                                {
                                    if (reader.Read())
                                    {
                                        if (reader["Total Valor"] != DBNull.Value)
                                        {
                                            string valorRealvalor = reader["Total Valor"].ToString();

                                            valorRealvalor = valorRealvalor.Trim();
                                            if (valorRealvalor.Contains(","))
                                            {
                                                valorRealvalor = valorRealvalor.Replace(",", ".");
                                            }

                                            if (double.TryParse(valorRealvalor, NumberStyles.Any, CultureInfo.InvariantCulture, out TotalValorReal))
                                            { }
                                            else
                                            {
                                                MessageBox.Show($"Erro ao converter o valor de 'Total Valor' para Número da Obra {NumeroObra}. Valor: {valorRealvalor}");
                                            }
                                        }
                                        else
                                        {
                                            MessageBox.Show($"Valor de 'Total Valor' não encontrado ou é nulo para o Número da Obra {NumeroObra}.");
                                        }
                                    }
                                    else
                                    {
                                        MessageBox.Show($"Número da obra {NumeroObra} não encontrado na tabela RealObras.");
                                    }
                                }
                            }

                            string queryOrcamentacaoValor = "SELECT [Total Valor] FROM dbo.Orçamentação WHERE [Numero da Obra] = @NumeroObra";
                            using (SqlCommand cmd = new SqlCommand(queryOrcamentacaoValor, BD.GetConnection()))
                            {
                                cmd.Parameters.AddWithValue("@NumeroObra", NumeroObra);

                                using (SqlDataReader reader = cmd.ExecuteReader())
                                {
                                    if (reader.Read())
                                    {
                                        if (reader["Total Valor"] != DBNull.Value)
                                        {
                                            string valorOrcamentoValor = reader["Total Valor"].ToString();

                                            valorOrcamentoValor = valorOrcamentoValor.Trim();
                                            if (valorOrcamentoValor.Contains(","))
                                            {
                                                valorOrcamentoValor = valorOrcamentoValor.Replace(",", ".");
                                            }

                                            if (double.TryParse(valorOrcamentoValor, NumberStyles.Any, CultureInfo.InvariantCulture, out TotalValorOrcamentacao))
                                            { }
                                            else
                                            {
                                                MessageBox.Show($"Erro ao converter o valor de 'Total Valor' para Número da Obra {NumeroObra}. Valor: {valorOrcamentoValor}");
                                            }
                                        }
                                        else
                                        {
                                            MessageBox.Show($"Valor de 'Total Valor' não encontrado ou é nulo para o Número da Obra {NumeroObra}.");
                                        }
                                    }
                                    else
                                    {
                                        MessageBox.Show($"Número da obra {NumeroObra} não encontrado na tabela Orçamentação.");
                                    }
                                }
                            }
                            TotalValorResultado = TotalValorReal - TotalValorOrcamentacao;
                            string updateQuery = "UPDATE dbo.ConclusaoObras " +
                                                 "SET [Total Valor] = @TotalValor " +
                                                 "WHERE [Numero da Obra] = @NumeroObra";

                            using (SqlCommand updateCmd = new SqlCommand(updateQuery, BD.GetConnection()))
                            {
                                updateCmd.Parameters.AddWithValue("@TotalValor", TotalValorResultado);
                                updateCmd.Parameters.AddWithValue("@NumeroObra", NumeroObra);

                                int rowsAffected = updateCmd.ExecuteNonQuery();
                            }
                        }
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
            else
            {
                MessageBox.Show("Não há dados na tabela.");
            }
        }
        private void AtualizarTabelaRealnaBdManual()
        {
            if (DataGridViewRealObras.Rows.Count > 0)
            {
                ComunicaBD BD = new ComunicaBD();
                try
                {
                    BD.ConectarBD();
                    foreach (DataGridViewRow row in DataGridViewRealObras.Rows)
                    {
                        if (DataGridViewRealObras.SelectedRows.Count > 0)
                        {
                            string NumeroObra = DataGridViewRealObras.SelectedRows[0].Cells["Numero da Obra"].Value.ToString();
                            string ID = DataGridViewRealObras.SelectedRows[0].Cells["ID"].Value.ToString();
                            string ValorEstrutura = DataGridViewRealObras.SelectedRows[0].Cells["Valor Estrutura"].Value.ToString();
                            string HorasEstrutura = DataGridViewRealObras.SelectedRows[0].Cells["Horas Estrutura"].Value.ToString();
                            string KgEstrutura = TextBoxKG.Text;
                            string KgEstrutura2 = SubstituirSeparadorDecimal(KgEstrutura);
                            ValorEstrutura = SubstituirSeparadorDecimal(ValorEstrutura);
                            HorasEstrutura = SubstituirSeparadorDecimal(HorasEstrutura);

                            if (string.IsNullOrEmpty(KgEstrutura))
                            {
                                MessageBox.Show("Erro: 'Kg Estrutura' não pode estar vazio.");
                                return;
                            }

                            double horasEstruturaDouble;
                            if (!double.TryParse(KgEstrutura2, NumberStyles.Any, CultureInfo.InvariantCulture, out horasEstruturaDouble))
                            {
                                MessageBox.Show("Erro: 'Horas Estrutura' não é um número válido.");
                                return;
                            }

                            double valorEstruturaDouble;
                            if (!double.TryParse(ValorEstrutura, NumberStyles.Any, CultureInfo.InvariantCulture, out valorEstruturaDouble))
                            {
                                MessageBox.Show("Erro: 'Valor Estrutura' não é um número válido.");
                                return;
                            }

                            double kgEstruturaDouble;
                            if (!double.TryParse(KgEstrutura, NumberStyles.Any, CultureInfo.InvariantCulture, out kgEstruturaDouble))
                            {
                                MessageBox.Show("Erro: 'Kg Estrutura' não é um número válido.");
                                return;
                            }

                            double kgEuroEstrutura = valorEstruturaDouble / kgEstruturaDouble;
                            double kgEuroEstruturaArredondado = Math.Round(kgEuroEstrutura, 3);
                            string kgEuroEstruturaString = kgEuroEstruturaArredondado.ToString("F3", CultureInfo.InvariantCulture);
                            MessageBox.Show($"{kgEuroEstruturaString}");
                            if (kgEuroEstruturaArredondado != 0)
                            {
                                kgEstruturaDouble = valorEstruturaDouble / kgEuroEstruturaArredondado;
                            }
                            else
                            {
                                MessageBox.Show("Erro: kgEuroEstrutura não pode ser zero.");
                                return;
                            }

                            string updateQuery = "UPDATE dbo.RealObras " +
                                "SET [Numero da Obra] = @NumeroObra, " +
                                "[KG Estrutura] = @KGestrutura, " +
                                "[KG/Euro Estrutura] = @KGEuroestrutura " +
                                "WHERE ID = @ID";

                            using (SqlCommand updateCmd = new SqlCommand(updateQuery, BD.GetConnection()))
                            {
                                updateCmd.Parameters.Add("@NumeroObra", SqlDbType.NVarChar).Value = NumeroObra;
                                updateCmd.Parameters.Add("@KGestrutura", SqlDbType.NVarChar).Value = KgEstrutura;
                                updateCmd.Parameters.Add("@KGEuroestrutura", SqlDbType.NVarChar).Value = kgEuroEstruturaString;
                                updateCmd.Parameters.Add("@ID", SqlDbType.Int).Value = Convert.ToInt32(ID);

                                int rowsAffected = updateCmd.ExecuteNonQuery();

                                if (rowsAffected > 0)
                                {
                                    MessageBox.Show("Dados registrados com sucesso!");
                                }
                                else
                                {
                                    MessageBox.Show("Nenhum dado foi atualizado.");
                                }
                            }
                        }
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
        private void CalcularTotaisEInserirNaBaseDeDados()
        {
            double totalKGEstrutura = 0;
            double totalHorasEstrutura = 0;
            double totalValorEstrutura = 0;
            double totalKGEuroEstrutura = 0;
            double totalHorasRevestimentos = 0;
            double totalValorRevestimentos = 0;
            double totalTotalHoras = 0;
            double totalTotalValor = 0;
            int linhaCountKGEuro = 0;
            PontoporvirgulaOrcamentacao();
            var numberFormatInfo = new System.Globalization.NumberFormatInfo();
            numberFormatInfo.CurrencyDecimalSeparator = ".";

            foreach (DataGridViewRow row in DataGridViewOrcamentacaoObras.Rows)
            {
                totalKGEstrutura += ProcessarValor(row.Cells["KG Estrutura"].Value);
                totalHorasEstrutura += ProcessarValor(row.Cells["Horas Estrutura"].Value);
                totalValorEstrutura += ProcessarValor(row.Cells["Valor Estrutura"].Value);
                totalHorasRevestimentos += ProcessarValor(row.Cells["Horas Revestimentos"].Value);
                totalValorRevestimentos += ProcessarValor(row.Cells["Valor Revestimentos"].Value);
                totalTotalHoras += ProcessarValor(row.Cells["Total Horas"].Value);
                totalTotalValor += ProcessarValor(row.Cells["Total Valor"].Value);

                double kgEuro = ProcessarValor(row.Cells["KG/Euro Estrutura"].Value);
                totalKGEuroEstrutura += kgEuro;
                if (kgEuro > 0) linhaCountKGEuro++;
            }

            double mediaKGEuroEstruturaReal = linhaCountKGEuro > 0 ? totalKGEuroEstrutura / linhaCountKGEuro : 0;
            int totalKGEstruturaInt = (int)Math.Round(totalKGEstrutura);
            int totalHorasEstruturaIntt = (int)Math.Round(totalHorasEstrutura);
            int totalHorasRevestimentosIntt = (int)Math.Round(totalHorasRevestimentos);
            int totalTotalHorasIntt = (int)Math.Round(totalTotalHoras);
            int totalValorEstruturaIntt = (int)Math.Round(totalValorEstrutura);
            int totalValorRevestimentosIntt = (int)Math.Round(totalValorRevestimentos);
            int totalTotalValorIntt = (int)Math.Round(totalTotalValor);

            string totalKGEstruturaRealInt = totalKGEstruturaInt.ToString() + " Kg";
            string totalHorasEstruturaInt = totalHorasEstruturaIntt.ToString() + " h";
            string totalKGEuroEstruturaInt = mediaKGEuroEstruturaReal.ToString("F2") + " €";
            string totalHorasRevestimentosInt = totalHorasRevestimentosIntt.ToString() + " h";
            string totalTotalHorasInt = totalTotalHorasIntt.ToString() + " h";

            string totalValorEstruturaInt = totalValorEstruturaIntt.ToString() + " €";
            string totalValorRevestimentosInt = totalValorRevestimentosIntt.ToString() + " €";
            string totalTotalValorInt = totalTotalValorIntt.ToString() + " €";

            ComunicaBD BD = new ComunicaBD();
            try
            {
                BD.ConectarBD();

                int obraId = 1;

                string query = "UPDATE dbo.TotalObras SET " +
                               "[Total KG Estrutura Orc] = @TotalKGEstrutura, " +
                               "[Total Horas Estrutura Orc] = @TotalHorasEstrutura, " +
                               "[Total Valor Estrutura Orc] = @TotalValorEstrutura, " +
                               "[Total KG/Euro Estrutura Orc] = @TotalKGEuroEstrutura, " +
                               "[Total Horas Revestimentos Orc] = @TotalHorasRevestimentos, " +
                               "[Total Valor Revestimentos Orc] = @TotalValorRevestimentos, " +
                               "[Total Horas Orc] = @TotalHoras, " +
                               "[Total Valor Orc] = @TotalValor " +
                               "WHERE ID = @ObraId";

                using (SqlCommand cmd = new SqlCommand(query, BD.GetConnection()))
                {
                    cmd.Parameters.AddWithValue("@TotalKGEstrutura", totalKGEstruturaRealInt);
                    cmd.Parameters.AddWithValue("@TotalHorasEstrutura", totalHorasEstruturaInt);
                    cmd.Parameters.AddWithValue("@TotalValorEstrutura", totalValorEstruturaInt);
                    cmd.Parameters.AddWithValue("@TotalKGEuroEstrutura", totalKGEuroEstruturaInt);
                    cmd.Parameters.AddWithValue("@TotalHorasRevestimentos", totalHorasRevestimentosInt);
                    cmd.Parameters.AddWithValue("@TotalValorRevestimentos", totalValorRevestimentosInt);
                    cmd.Parameters.AddWithValue("@TotalHoras", totalTotalHorasInt);
                    cmd.Parameters.AddWithValue("@TotalValor", totalTotalValorInt);
                    cmd.Parameters.AddWithValue("@ObraId", obraId);

                    cmd.ExecuteNonQuery();
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
        private void CalcularTotaisReal()
        {
            double totalKGEstruturaReal = 0;
            double totalHorasEstruturaReal = 0;
            double totalValorEstruturaReal = 0;
            double totalKGEuroEstruturaReal = 0;
            double totalHorasRevestimentosReal = 0;
            double totalValorRevestimentosReal = 0;
            double totalHorasAprovacaoReal = 0;
            double totalValorAprovacaoReal = 0;
            double totalHorasAlteracoesReal = 0;
            double totalValorAlteracoesReal = 0;
            double totalHorasFabricoReal = 0;
            double totalValorFabricoReal = 0;
            double totalHorasSoldaduraReal = 0;
            double totalValorSoldaduraReal = 0;
            double totalHorasMontagemReal = 0;
            double totalValorMontagemReal = 0;
            double totalHorasDiversosReal = 0;
            double totalValorDiversosReal = 0;
            double totalHorasReal = 0;
            double totalValorReal = 0;
            int linhaCountKGEuro = 0;

            PontoporvirgulaReal();

            foreach (DataGridViewRow row in DataGridViewRealObras.Rows)
            {
                totalKGEstruturaReal += ProcessarValor(row.Cells["KG Estrutura"].Value);
                totalHorasEstruturaReal += ProcessarValor(row.Cells["Horas Estrutura"].Value);
                totalValorEstruturaReal += ProcessarValor(row.Cells["Valor Estrutura"].Value);

                double kgEuro = ProcessarValor(row.Cells["KG/Euro Estrutura"].Value);
                totalKGEuroEstruturaReal += kgEuro;
                if (kgEuro > 0) linhaCountKGEuro++;

                totalHorasRevestimentosReal += ProcessarValor(row.Cells["Horas Revestimentos"].Value);
                totalValorRevestimentosReal += ProcessarValor(row.Cells["Valor Revestimentos"].Value);

                totalHorasAprovacaoReal += ProcessarValor(row.Cells["Horas Aprovação"].Value);
                totalValorAprovacaoReal += ProcessarValor(row.Cells["Valor Aprovação"].Value);

                totalHorasAlteracoesReal += ProcessarValor(row.Cells["Horas Alterações"].Value);
                totalValorAlteracoesReal += ProcessarValor(row.Cells["Valor Alterações"].Value);

                totalHorasFabricoReal += ProcessarValor(row.Cells["Horas Fabrico"].Value);
                totalValorFabricoReal += ProcessarValor(row.Cells["Valor Fabrico"].Value);

                totalHorasSoldaduraReal += ProcessarValor(row.Cells["Horas Soldadura"].Value);
                totalValorSoldaduraReal += ProcessarValor(row.Cells["Valor Soldadura"].Value);

                totalHorasMontagemReal += ProcessarValor(row.Cells["Horas Montagem"].Value);
                totalValorMontagemReal += ProcessarValor(row.Cells["Valor Montagem"].Value);

                totalHorasDiversosReal += ProcessarValor(row.Cells["Horas Diversos"].Value);
                totalValorDiversosReal += ProcessarValor(row.Cells["Valor Diversos"].Value);

                totalHorasReal += ProcessarValor(row.Cells["Total Horas"].Value);
                totalValorReal += ProcessarValor(row.Cells["Total Valor"].Value);
            }

            double mediaKGEuroEstruturaReal = linhaCountKGEuro > 0 ? totalKGEuroEstruturaReal / linhaCountKGEuro : 0;
            int totalKGEstruturaRealIntt = (int)Math.Round(totalKGEstruturaReal);
            int totalHorasEstruturaRealIntt = (int)Math.Round(totalHorasEstruturaReal);
            int totalHorasRevestimentosRealIntt = (int)Math.Round(totalHorasRevestimentosReal);
            int totalHorasAprovacaoRealIntt = (int)Math.Round(totalHorasAprovacaoReal);
            int totalHorasAlteracoesRealIntt = (int)Math.Round(totalHorasAlteracoesReal);
            int totalHorasFabricoRealIntt = (int)Math.Round(totalHorasFabricoReal);
            int totalHorasSoldaduraRealIntt = (int)Math.Round(totalHorasSoldaduraReal);
            int totalHorasMontagemRealIntt = (int)Math.Round(totalHorasMontagemReal);
            int totalHorasDiversosRealIntt = (int)Math.Round(totalHorasDiversosReal);
            int totalHorasRealIntt = (int)Math.Round(totalHorasReal);


            string totalKGEstruturaRealInt = totalKGEstruturaRealIntt.ToString() + " Kg";
            string totalHorasEstruturaRealInt = totalHorasEstruturaRealIntt.ToString() + " h";
            string totalHorasRevestimentosRealInt = totalHorasRevestimentosRealIntt.ToString() + " h";
            string totalHorasAprovacaoRealInt = totalHorasAprovacaoRealIntt.ToString() + " h";
            string totalHorasAlteracoesRealInt = totalHorasAlteracoesRealIntt.ToString() + " h";
            string totalHorasFabricoRealInt = totalHorasFabricoRealIntt.ToString() + " h";
            string totalHorasSoldaduraRealInt = totalHorasSoldaduraRealIntt.ToString() + " h";
            string totalHorasMontagemRealInt = totalHorasMontagemRealIntt.ToString() + " h";
            string totalHorasDiversosRealInt = totalHorasDiversosRealIntt.ToString() + " h";
            string totalHorasRealInt = totalHorasRealIntt.ToString() + " h";

            string totalValorEstruturaReall = totalValorEstruturaReal.ToString() + " €";
            string totalKGEuroEstruturaReall = mediaKGEuroEstruturaReal.ToString("F2") + " kg";
            string totalValorRevestimentosReall = totalValorRevestimentosReal.ToString() + " €";
            string totalValorAprovacaoReall = totalValorAprovacaoReal.ToString() + " €";
            string totalValorAlteracoesReall = totalValorAlteracoesReal.ToString() + " €";
            string totalValorFabricoReall = totalValorFabricoReal.ToString() + " €";
            string totalValorSoldaduraReall = totalValorSoldaduraReal.ToString() + " €";
            string totalValorMontagemReall = totalValorMontagemReal.ToString() + " €";
            string totalValorDiversosReall = totalValorDiversosReal.ToString() + " €";
            string totalValorReall = totalValorReal.ToString() + " €";

            ComunicaBD BD = new ComunicaBD();

            try
            {
                BD.ConectarBD();

                int obraId = 1;

                string query = "UPDATE dbo.TotalObras SET " +
                                "[Total KG Estrutura Real] = @TotalKGEstruturaReal, " +
                                "[Total Horas Estrutura Real] = @TotalHorasEstruturaReal, " +
                                "[Total Valor Estrutura Real] = @TotalValorEstruturaReal, " +
                                "[Total KG/Euro Estrutura Real] = @TotalKGEuroEstruturaReal, " +
                                "[Total Horas Revestimentos Real] = @TotalHorasRevestimentosReal, " +
                                "[Total Valor Revestimentos Real] = @TotalValorRevestimentosReal, " +
                                "[Total Horas Aprovacao Real] = @TotalHorasAprovacaoReal, " +
                                "[Total Valor Aprovacao Real] = @TotalValorAprovacaoReal, " +
                                "[Total Horas Alteracoes Real] = @TotalHorasAlteracoesReal, " +
                                "[Total Valor Alteracoes Real] = @TotalValorAlteracoesReal, " +
                                "[Total Horas Fabrico Real] = @TotalHorasFabricoReal, " +
                                "[Total Valor Fabrico Real] = @TotalValorFabricoReal, " +
                                "[Total Horas Soldadura Real] = @TotalHorasSoldaduraReal, " +
                                "[Total Valor Soldadura Real] = @TotalValorSoldaduraReal, " +
                                "[Total Horas Montagem Real] = @TotalHorasMontagemReal, " +
                                "[Total Valor Montagem Real] = @TotalValorMontagemReal, " +
                                "[Total Horas Diversos Real] = @TotalHorasDiversosReal, " +
                                "[Total Valor Diversos Real] = @TotalValorDiversosReal, " +
                                "[Total Horas Real] = @TotalHorasReal, " +
                                "[Total Valor Real] = @TotalValorReal " +
                                "WHERE ID = @ObraId";

                using (SqlCommand cmd = new SqlCommand(query, BD.GetConnection()))
                {
                    cmd.Parameters.AddWithValue("@TotalKGEstruturaReal", totalKGEstruturaRealInt);
                    cmd.Parameters.AddWithValue("@TotalHorasEstruturaReal", totalHorasEstruturaRealInt);
                    cmd.Parameters.AddWithValue("@TotalValorEstruturaReal", totalValorEstruturaReall);
                    cmd.Parameters.AddWithValue("@TotalKGEuroEstruturaReal", totalKGEuroEstruturaReall);
                    cmd.Parameters.AddWithValue("@TotalHorasRevestimentosReal", totalHorasRevestimentosRealInt);
                    cmd.Parameters.AddWithValue("@TotalValorRevestimentosReal", totalValorRevestimentosReall);
                    cmd.Parameters.AddWithValue("@TotalHorasAprovacaoReal", totalHorasAprovacaoRealInt);
                    cmd.Parameters.AddWithValue("@TotalValorAprovacaoReal", totalValorAprovacaoReall);
                    cmd.Parameters.AddWithValue("@TotalHorasAlteracoesReal", totalHorasAlteracoesRealInt);
                    cmd.Parameters.AddWithValue("@TotalValorAlteracoesReal", totalValorAlteracoesReall);
                    cmd.Parameters.AddWithValue("@TotalHorasFabricoReal", totalHorasFabricoRealInt);
                    cmd.Parameters.AddWithValue("@TotalValorFabricoReal", totalValorFabricoReall);
                    cmd.Parameters.AddWithValue("@TotalHorasSoldaduraReal", totalHorasSoldaduraRealInt);
                    cmd.Parameters.AddWithValue("@TotalValorSoldaduraReal", totalValorSoldaduraReall);
                    cmd.Parameters.AddWithValue("@TotalHorasMontagemReal", totalHorasMontagemRealInt);
                    cmd.Parameters.AddWithValue("@TotalValorMontagemReal", totalValorMontagemReall);
                    cmd.Parameters.AddWithValue("@TotalHorasDiversosReal", totalHorasDiversosRealInt);
                    cmd.Parameters.AddWithValue("@TotalValorDiversosReal", totalValorDiversosReall);
                    cmd.Parameters.AddWithValue("@TotalHorasReal", totalHorasRealInt);
                    cmd.Parameters.AddWithValue("@TotalValorReal", totalValorReall);
                    cmd.Parameters.AddWithValue("@ObraId", obraId);
                    cmd.ExecuteNonQuery();
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
        private void CalcularTotaisRealPercentagem()
        {
            double totalPercentagemEstruturaReal = 0;
            double totalPercentagemRevestimentosReal = 0;
            double totalPercentagemAprovacaoReal = 0;
            double totalPercentagemAlteracoesReal = 0;
            double totalPercentagemFabricoReal = 0;
            double totalPercentagemSoldaduraReal = 0;
            double totalPercentagemMontagemReal = 0;
            double totalPercentagemDiversosReal = 0;
            int linhaCount = 0;
            PontoporvirgulaReal();

            foreach (DataGridViewRow row in DataGridViewRealObras.Rows)
            {
                if (row.Cells["Percentagem Estrutura"].Value != null)
                {
                    totalPercentagemEstruturaReal += CalcularPercentagem(row.Cells["Percentagem Estrutura"].Value.ToString());
                    linhaCount++;
                }
                if (row.Cells["Percentagem Revestimentos"].Value != null)
                {
                    totalPercentagemRevestimentosReal += CalcularPercentagem(row.Cells["Percentagem Revestimentos"].Value.ToString());
                }
                if (row.Cells["Percentagem Aprovação"].Value != null)
                {
                    totalPercentagemAprovacaoReal += CalcularPercentagem(row.Cells["Percentagem Aprovação"].Value.ToString());
                }
                if (row.Cells["Percentagem Alterações"].Value != null)
                {
                    totalPercentagemAlteracoesReal += CalcularPercentagem(row.Cells["Percentagem Alterações"].Value.ToString());
                }
                if (row.Cells["Percentagem Fabrico"].Value != null)
                {
                    totalPercentagemFabricoReal += CalcularPercentagem(row.Cells["Percentagem Fabrico"].Value.ToString());
                }
                if (row.Cells["Percentagem Soldadura"].Value != null)
                {
                    totalPercentagemSoldaduraReal += CalcularPercentagem(row.Cells["Percentagem Soldadura"].Value.ToString());
                }
                if (row.Cells["Percentagem Montagem"].Value != null)
                {
                    totalPercentagemMontagemReal += CalcularPercentagem(row.Cells["Percentagem Montagem"].Value.ToString());
                }
                if (row.Cells["Percentagem Diversos"].Value != null)
                {
                    totalPercentagemDiversosReal += CalcularPercentagem(row.Cells["Percentagem Diversos"].Value.ToString());
                }
            }
            double mediaPercentagemEstruturaReal = linhaCount > 0 ? totalPercentagemEstruturaReal / linhaCount : 0;
            double mediaPercentagemRevestimentosReal = linhaCount > 0 ? totalPercentagemRevestimentosReal / linhaCount : 0;
            double mediaPercentagemAprovacaoReal = linhaCount > 0 ? totalPercentagemAprovacaoReal / linhaCount : 0;
            double mediaPercentagemAlteracoesReal = linhaCount > 0 ? totalPercentagemAlteracoesReal / linhaCount : 0;
            double mediaPercentagemFabricoReal = linhaCount > 0 ? totalPercentagemFabricoReal / linhaCount : 0;
            double mediaPercentagemSoldaduraReal = linhaCount > 0 ? totalPercentagemSoldaduraReal / linhaCount : 0;
            double mediaPercentagemMontagemReal = linhaCount > 0 ? totalPercentagemMontagemReal / linhaCount : 0;
            double mediaPercentagemDiversosReal = linhaCount > 0 ? totalPercentagemDiversosReal / linhaCount : 0;

            string percentualEstruturaReal = mediaPercentagemEstruturaReal.ToString("F2") + " %";
            string percentualRevestimentosReal = mediaPercentagemRevestimentosReal.ToString("F2") + " %";
            string percentualAprovacaoReal = mediaPercentagemAprovacaoReal.ToString("F2") + " %";
            string percentualAlteracoesReal = mediaPercentagemAlteracoesReal.ToString("F2") + " %";
            string percentualFabricoReal = mediaPercentagemFabricoReal.ToString("F2") + " %";
            string percentualSoldaduraReal = mediaPercentagemSoldaduraReal.ToString("F2") + " %";
            string percentualMontagemReal = mediaPercentagemMontagemReal.ToString("F2") + " %";
            string percentualDiversosReal = mediaPercentagemDiversosReal.ToString("F2") + " %";

            ComunicaBD BD = new ComunicaBD();
            try
            {
                BD.ConectarBD();

                int obraId = 1;

                string query = "UPDATE dbo.TotalObras SET " +
                                 "[Percentagem Estrutura Real] = @PercentagemEstruturaReal, " +
                                 "[Percentagem Revestimentos Real] = @PercentagemRevestimentosReal, " +
                                 "[Percentagem Aprovacao Real] = @PercentagemAprovacaoReal, " +
                                 "[Percentagem Alteracoes Real] = @PercentagemAlteracoesReal, " +
                                 "[Percentagem Fabrico Real] = @PercentagemFabricoReal, " +
                                 "[Percentagem Soldadura Real] = @PercentagemSoldaduraReal, " +
                                 "[Percentagem Montagem Real] = @PercentagemMontagemReal, " +
                                 "[Percentagem Diversos Real] = @PercentagemDiversosReal " +
                                 "WHERE ID = @ObraId";

                using (SqlCommand cmd = new SqlCommand(query, BD.GetConnection()))
                {
                    cmd.Parameters.AddWithValue("@PercentagemEstruturaReal", percentualEstruturaReal);
                    cmd.Parameters.AddWithValue("@PercentagemRevestimentosReal", percentualRevestimentosReal);
                    cmd.Parameters.AddWithValue("@PercentagemAprovacaoReal", percentualAprovacaoReal);
                    cmd.Parameters.AddWithValue("@PercentagemAlteracoesReal", percentualAlteracoesReal);
                    cmd.Parameters.AddWithValue("@PercentagemFabricoReal", percentualFabricoReal);
                    cmd.Parameters.AddWithValue("@PercentagemSoldaduraReal", percentualSoldaduraReal);
                    cmd.Parameters.AddWithValue("@PercentagemMontagemReal", percentualMontagemReal);
                    cmd.Parameters.AddWithValue("@PercentagemDiversosReal", percentualDiversosReal);
                    cmd.Parameters.AddWithValue("@ObraId", obraId);
                    cmd.ExecuteNonQuery();
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
        public void CarregarTipologiaNaComboBox()
        {
            ComunicaBD BD = new ComunicaBD();
            BD.ConectarBD();
            string query = "SELECT Tipologia FROM dbo.Tipologia";
            List<string> list = BD.Procurarbdlist(query);
            BD.DesonectarBD();
            ComboBoxTipologiaInserir.Items.Clear();
            ComboBoxTipologiaFiltro.Items.Clear();
            if (list.Count > 0)
            {
                foreach (string nome in list)
                {
                    ComboBoxTipologiaInserir.Items.Add(nome);
                    ComboBoxTipologiaFiltro.Items.Add(nome);
                }
            }
            else
            {
                MessageBox.Show("Nenhuma Tipologia encontrado na base de dados.");
            }
        }
        private void ConfigurarComboBoxTipologia()
        {
            if (DataGridViewOrcamentacaoObras.Columns["Tipologia"] == null)
            {
                DataGridViewComboBoxColumn comboBoxColumn = new DataGridViewComboBoxColumn
                {
                    Name = "Tipologia",
                    HeaderText = "Tipologia",
                    DataPropertyName = "Tipologia",
                    FlatStyle = FlatStyle.Popup
                };
                comboBoxColumn.Items.AddRange("Porticado", "Treliçado", "Revestimentos");
                DataGridViewOrcamentacaoObras.Columns.Add(comboBoxColumn);
            }
        }
        private void ConfigurarComboBoxPreparadorResponsavel()
        {
            if (DataGridViewOrcamentacaoObras.Columns["Preparador Responsavel"] == null)
            {
                DataGridViewComboBoxColumn comboBoxColumn = new DataGridViewComboBoxColumn
                {
                    Name = "Preparador Responsavel",
                    HeaderText = "Preparador Responsavel",
                    DataPropertyName = "PreparadorResponsavel",
                    FlatStyle = FlatStyle.Popup
                };
                foreach (var item in ComboBoxPreparadorAdd.Items)
                {
                    comboBoxColumn.Items.Add(item);
                }
                DataGridViewOrcamentacaoObras.Columns.Add(comboBoxColumn);
            }
        }
        private void PontoporvirgulaReal()
        {
            foreach (DataGridViewRow row in DataGridViewRealObras.Rows)
            {
                foreach (DataGridViewCell cell in row.Cells)
                {
                    if (cell.Value != null && cell.Value is string)
                    {
                        string cellValue = cell.Value.ToString();

                        cell.Value = cellValue.Replace(",", ".");
                    }
                }
            }
        }
        private void CalcularTotaisConclusao()
        {
            double totalPercentagemTotal = 0;
            int linhaCount = 0;
            PontoporvirgulaReal();
            foreach (DataGridViewRow row in DataGridViewConclusaoObras.Rows)
            {
                if (row.Cells["Percentagem Total"].Value != null)
                {
                    totalPercentagemTotal += CalcularPercentagem(row.Cells["Percentagem Total"].Value.ToString());
                    linhaCount++;
                }
            }
            double mediaPercentagemTotal = linhaCount > 0 ? totalPercentagemTotal / linhaCount : 0;
            string percentualComSimbolo = mediaPercentagemTotal.ToString("F2") + " %";
            double totalHoras = 0;
            double totalValor = 0;

            foreach (DataGridViewRow row in DataGridViewConclusaoObras.Rows)
            {
                if (row.Cells["Total Horas"].Value != null)
                {
                    totalHoras += ProcessarValor(row.Cells["Total Horas"].Value.ToString());
                }
                if (row.Cells["Total Valor"].Value != null)
                {
                    totalValor += ProcessarValor(row.Cells["Total Valor"].Value.ToString());
                }
            }
            string totalValorr = totalValor.ToString() + " €";
            string totalHorass = totalHoras.ToString() + " h";

            ComunicaBD BD = new ComunicaBD();
            try
            {
                BD.ConectarBD();

                int obraId = 1;

                string query = "UPDATE dbo.TotalObras SET " +
                                 "[Total Horas Concl] = @TotalHoras, " +
                                 "[Total Valor Concl] = @TotalValororas, " +
                                 "[Percentagem Total Concl] = @PercentagemTotal " +
                                 "WHERE ID = @ObraId";

                using (SqlCommand cmd = new SqlCommand(query, BD.GetConnection()))
                {
                    cmd.Parameters.AddWithValue("@TotalHoras", totalHorass);
                    cmd.Parameters.AddWithValue("@TotalValororas", totalValorr);
                    cmd.Parameters.AddWithValue("@PercentagemTotal", percentualComSimbolo);
                    cmd.Parameters.AddWithValue("@ObraId", obraId);

                    cmd.ExecuteNonQuery();
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
        public void CarregarPastasNaComboBoxAno()
        {
            string caminhoPasta = @"\\marconi\COMPANY SHARED FOLDER\OFELIZ\OFM\2.AN\2.CM\DP\1 Obras";
            if (!Directory.Exists(caminhoPasta))
            {
                MessageBox.Show("O caminho especificado não existe.");
                return;
            }
            string[] subpastas = Directory.GetDirectories(caminhoPasta);
            ComboBoxAnoAdd.Items.Clear();
            ComboBoxAnoAdd2.Items.Clear();
            foreach (string subpasta in subpastas)
            {
                string nomePasta = Path.GetFileName(subpasta);

                if (int.TryParse(nomePasta, out int ano))
                {
                    if (ano >= 2020)
                    {
                        ComboBoxAnoAdd.Items.Add(nomePasta);
                        ComboBoxAnoAdd2.Items.Add(nomePasta);
                    }
                }
            }
        }        
        private void AdicionarSufixosNasColunas()
        {
            foreach (DataGridViewRow row in DataGridViewOrcamentacaoObras.Rows)
            {
                if (row.Cells["KG Estrutura"].Value != DBNull.Value && row.Cells["KG Estrutura"].Value != null)
                {
                    double kgValue;
                    bool isKgNumeric = Double.TryParse(row.Cells["KG Estrutura"].Value.ToString(), out kgValue);

                    if (isKgNumeric)
                    {
                        row.Cells["KG Estrutura"].Value = $"{kgValue} kg";
                    }
                }
                if (row.Cells["Horas Estrutura"].Value != DBNull.Value && row.Cells["Horas Estrutura"].Value != null)
                {
                    double horasValue;
                    bool isHorasNumeric = Double.TryParse(row.Cells["Horas Estrutura"].Value.ToString(), out horasValue);

                    if (isHorasNumeric)
                    {
                        row.Cells["Horas Estrutura"].Value = $"{horasValue} h";
                    }
                }
                if (row.Cells["Valor Estrutura"].Value != DBNull.Value && row.Cells["Valor Estrutura"].Value != null)
                {
                    string valorStr = row.Cells["Valor Estrutura"].Value.ToString();

                    valorStr = valorStr.Replace('.', ',');

                    double valorValue;
                    bool isValorNumeric = Double.TryParse(valorStr, out valorValue);

                    if (isValorNumeric)
                    {
                        row.Cells["Valor Estrutura"].Value = $"{valorValue:F2} €";
                    }
                }
                if (row.Cells["KG/Euro Estrutura"].Value != DBNull.Value && row.Cells["KG/Euro Estrutura"].Value != null)
                {
                    string valorStr = row.Cells["KG/Euro Estrutura"].Value.ToString();

                    valorStr = valorStr.Replace('.', ',');

                    double valorValue;
                    bool isValorNumeric = Double.TryParse(valorStr, out valorValue);

                    if (isValorNumeric)
                    {
                        row.Cells["KG/Euro Estrutura"].Value = $"{valorValue:F2} €";
                    }
                }
                if (row.Cells["Horas Revestimentos"].Value != DBNull.Value && row.Cells["Horas Revestimentos"].Value != null)
                {
                    double horasValue;
                    bool isHorasNumeric = Double.TryParse(row.Cells["Horas Revestimentos"].Value.ToString(), out horasValue);

                    if (isHorasNumeric)
                    {
                        row.Cells["Horas Revestimentos"].Value = $"{horasValue} h";
                    }
                }
                if (row.Cells["Valor Revestimentos"].Value != DBNull.Value && row.Cells["Valor Revestimentos"].Value != null)
                {
                    string valorStr = row.Cells["Valor Revestimentos"].Value.ToString();

                    valorStr = valorStr.Replace('.', ',');

                    double valorValue;
                    bool isValorNumeric = Double.TryParse(valorStr, out valorValue);

                    if (isValorNumeric)
                    {
                        row.Cells["Valor Revestimentos"].Value = $"{valorValue:F2} €";
                    }
                }
                if (row.Cells["Total Horas"].Value != DBNull.Value && row.Cells["Total Horas"].Value != null)
                {
                    double horasValue;
                    bool isHorasNumeric = Double.TryParse(row.Cells["Total Horas"].Value.ToString(), out horasValue);

                    if (isHorasNumeric)
                    {
                        row.Cells["Total Horas"].Value = $"{horasValue} h";
                    }
                }
                if (row.Cells["Total Valor"].Value != DBNull.Value && row.Cells["Total Valor"].Value != null)
                {
                    string valorStr = row.Cells["Total Valor"].Value.ToString();

                    valorStr = valorStr.Replace('.', ',');

                    double valorValue;
                    bool isValorNumeric = Double.TryParse(valorStr, out valorValue);

                    if (isValorNumeric)
                    {
                        row.Cells["Total Valor"].Value = $"{valorValue:F2} €";
                    }
                }

            }

            foreach (DataGridViewRow row in DataGridViewRealObras.Rows)
            {
                if (row.Cells["KG Estrutura"].Value != DBNull.Value && row.Cells["KG Estrutura"].Value != null)
                {
                    double kgValue;
                    bool isKgNumeric = Double.TryParse(row.Cells["KG Estrutura"].Value.ToString(), out kgValue);

                    if (isKgNumeric)
                    {
                        row.Cells["KG Estrutura"].Value = $"{kgValue} kg";
                    }
                }
                if (row.Cells["Horas Estrutura"].Value != DBNull.Value && row.Cells["Horas Estrutura"].Value != null)
                {
                    double horasValue;
                    bool isHorasNumeric = Double.TryParse(row.Cells["Horas Estrutura"].Value.ToString(), out horasValue);

                    if (isHorasNumeric)
                    {
                        row.Cells["Horas Estrutura"].Value = $"{horasValue} h";
                    }
                }
                if (row.Cells["Valor Estrutura"].Value != DBNull.Value && row.Cells["Valor Estrutura"].Value != null)
                {
                    string valorStr = row.Cells["Valor Estrutura"].Value.ToString();

                    valorStr = valorStr.Replace('.', ',');

                    double valorValue;
                    bool isValorNumeric = Double.TryParse(valorStr, out valorValue);

                    if (isValorNumeric)
                    {
                        row.Cells["Valor Estrutura"].Value = $"{valorValue:F2} €";
                    }
                }
                if (row.Cells["KG/Euro Estrutura"].Value != DBNull.Value && row.Cells["KG/Euro Estrutura"].Value != null)
                {
                    string valorStr = row.Cells["KG/Euro Estrutura"].Value.ToString();

                    valorStr = valorStr.Replace('.', ',');

                    double valorValue;
                    bool isValorNumeric = Double.TryParse(valorStr, out valorValue);

                    if (isValorNumeric)
                    {
                        row.Cells["KG/Euro Estrutura"].Value = $"{valorValue:F2} €";
                    }
                }
                if (row.Cells["Horas Revestimentos"].Value != DBNull.Value && row.Cells["Horas Revestimentos"].Value != null)
                {
                    double horasValue;
                    bool isHorasNumeric = Double.TryParse(row.Cells["Horas Revestimentos"].Value.ToString(), out horasValue);

                    if (isHorasNumeric)
                    {
                        row.Cells["Horas Revestimentos"].Value = $"{horasValue} h";
                    }
                }
                if (row.Cells["Valor Revestimentos"].Value != DBNull.Value && row.Cells["Valor Revestimentos"].Value != null)
                {
                    string valorStr = row.Cells["Valor Revestimentos"].Value.ToString();

                    valorStr = valorStr.Replace('.', ',');

                    double valorValue;
                    bool isValorNumeric = Double.TryParse(valorStr, out valorValue);

                    if (isValorNumeric)
                    {
                        row.Cells["Valor Revestimentos"].Value = $"{valorValue:F2} €";
                    }
                }
                if (row.Cells["Horas Aprovação"].Value != DBNull.Value && row.Cells["Horas Aprovação"].Value != null)
                {
                    double horasValue;
                    bool isHorasNumeric = Double.TryParse(row.Cells["Horas Aprovação"].Value.ToString(), out horasValue);

                    if (isHorasNumeric)
                    {
                        row.Cells["Horas Aprovação"].Value = $"{horasValue} h";
                    }
                }
                if (row.Cells["Valor Aprovação"].Value != DBNull.Value && row.Cells["Valor Aprovação"].Value != null)
                {
                    string valorStr = row.Cells["Valor Aprovação"].Value.ToString();

                    valorStr = valorStr.Replace('.', ',');

                    double valorValue;
                    bool isValorNumeric = Double.TryParse(valorStr, out valorValue);

                    if (isValorNumeric)
                    {
                        row.Cells["Valor Aprovação"].Value = $"{valorValue:F2} €";
                    }
                }
                if (row.Cells["Horas Alterações"].Value != DBNull.Value && row.Cells["Horas Alterações"].Value != null)
                {
                    double horasValue;
                    bool isHorasNumeric = Double.TryParse(row.Cells["Horas Alterações"].Value.ToString(), out horasValue);

                    if (isHorasNumeric)
                    {
                        row.Cells["Horas Alterações"].Value = $"{horasValue} h";
                    }
                }
                if (row.Cells["Valor Alterações"].Value != DBNull.Value && row.Cells["Valor Alterações"].Value != null)
                {
                    string valorStr = row.Cells["Valor Alterações"].Value.ToString();

                    valorStr = valorStr.Replace('.', ',');

                    double valorValue;
                    bool isValorNumeric = Double.TryParse(valorStr, out valorValue);

                    if (isValorNumeric)
                    {
                        row.Cells["Valor Alterações"].Value = $"{valorValue:F2} €";
                    }
                }
                if (row.Cells["Horas Fabrico"].Value != DBNull.Value && row.Cells["Horas Fabrico"].Value != null)
                {
                    double horasValue;
                    bool isHorasNumeric = Double.TryParse(row.Cells["Horas Fabrico"].Value.ToString(), out horasValue);

                    if (isHorasNumeric)
                    {
                        row.Cells["Horas Fabrico"].Value = $"{horasValue} h";
                    }
                }
                if (row.Cells["Valor Fabrico"].Value != DBNull.Value && row.Cells["Valor Fabrico"].Value != null)
                {
                    string valorStr = row.Cells["Valor Fabrico"].Value.ToString();

                    valorStr = valorStr.Replace('.', ',');

                    double valorValue;
                    bool isValorNumeric = Double.TryParse(valorStr, out valorValue);

                    if (isValorNumeric)
                    {
                        row.Cells["Valor Fabrico"].Value = $"{valorValue:F2} €";
                    }
                }
                if (row.Cells["Horas Soldadura"].Value != DBNull.Value && row.Cells["Horas Soldadura"].Value != null)
                {
                    double horasValue;
                    bool isHorasNumeric = Double.TryParse(row.Cells["Horas Soldadura"].Value.ToString(), out horasValue);

                    if (isHorasNumeric)
                    {
                        row.Cells["Horas Soldadura"].Value = $"{horasValue} h";
                    }
                }
                if (row.Cells["Valor Soldadura"].Value != DBNull.Value && row.Cells["Valor Soldadura"].Value != null)
                {
                    string valorStr = row.Cells["Valor Soldadura"].Value.ToString();

                    valorStr = valorStr.Replace('.', ',');

                    double valorValue;
                    bool isValorNumeric = Double.TryParse(valorStr, out valorValue);

                    if (isValorNumeric)
                    {
                        row.Cells["Valor Soldadura"].Value = $"{valorValue:F2} €";
                    }
                }
                if (row.Cells["Horas Montagem"].Value != DBNull.Value && row.Cells["Horas Montagem"].Value != null)
                {
                    double horasValue;
                    bool isHorasNumeric = Double.TryParse(row.Cells["Horas Montagem"].Value.ToString(), out horasValue);

                    if (isHorasNumeric)
                    {
                        row.Cells["Horas Montagem"].Value = $"{horasValue} h";
                    }
                }
                if (row.Cells["Valor Montagem"].Value != DBNull.Value && row.Cells["Valor Montagem"].Value != null)
                {
                    string valorStr = row.Cells["Valor Montagem"].Value.ToString();

                    valorStr = valorStr.Replace('.', ',');

                    double valorValue;
                    bool isValorNumeric = Double.TryParse(valorStr, out valorValue);

                    if (isValorNumeric)
                    {
                        row.Cells["Valor Montagem"].Value = $"{valorValue:F2} €";
                    }
                }
                if (row.Cells["Horas Diversos"].Value != DBNull.Value && row.Cells["Horas Diversos"].Value != null)
                {
                    double horasValue;
                    bool isHorasNumeric = Double.TryParse(row.Cells["Horas Diversos"].Value.ToString(), out horasValue);

                    if (isHorasNumeric)
                    {
                        row.Cells["Horas Diversos"].Value = $"{horasValue} h";
                    }
                }
                if (row.Cells["Valor Diversos"].Value != DBNull.Value && row.Cells["Valor Diversos"].Value != null)
                {
                    string valorStr = row.Cells["Valor Diversos"].Value.ToString();

                    valorStr = valorStr.Replace('.', ',');

                    double valorValue;
                    bool isValorNumeric = Double.TryParse(valorStr, out valorValue);

                    if (isValorNumeric)
                    {
                        row.Cells["Valor Diversos"].Value = $"{valorValue:F2} €";
                    }
                }
                if (row.Cells["Total Horas"].Value != DBNull.Value && row.Cells["Total Horas"].Value != null)
                {
                    double horasValue;
                    bool isHorasNumeric = Double.TryParse(row.Cells["Total Horas"].Value.ToString(), out horasValue);

                    if (isHorasNumeric)
                    {
                        row.Cells["Total Horas"].Value = $"{horasValue} h";
                    }
                }
                if (row.Cells["Total Valor"].Value != DBNull.Value && row.Cells["Total Valor"].Value != null)
                {
                    string valorStr = row.Cells["Total Valor"].Value.ToString();

                    valorStr = valorStr.Replace('.', ',');

                    double valorValue;
                    bool isValorNumeric = Double.TryParse(valorStr, out valorValue);

                    if (isValorNumeric)
                    {
                        row.Cells["Total Valor"].Value = $"{valorValue:F2} €";
                    }
                }
            }

            foreach (DataGridViewRow row in DataGridViewConclusaoObras.Rows)
            {
                if (row.Cells["Total Horas"].Value != DBNull.Value && row.Cells["Total Horas"].Value != null)
                {
                    double horasValue;
                    bool isHorasNumeric = Double.TryParse(row.Cells["Total Horas"].Value.ToString(), out horasValue);

                    if (isHorasNumeric)
                    {
                        row.Cells["Total Horas"].Value = $"{horasValue} h";
                    }
                }
                if (row.Cells["Total Valor"].Value != DBNull.Value && row.Cells["Total Valor"].Value != null)
                {
                    string valorStr = row.Cells["Total Valor"].Value.ToString();

                    valorStr = valorStr.Replace('.', ',');

                    double valorValue;
                    bool isValorNumeric = Double.TryParse(valorStr, out valorValue);

                    if (isValorNumeric)
                    {
                        row.Cells["Total Valor"].Value = $"{valorValue:F2} €";
                    }
                }
                if (row.Cells["Dias de Preparação"].Value != DBNull.Value && row.Cells["Dias de Preparação"].Value != null)
                {
                    double horasValue;
                    bool isHorasNumeric = Double.TryParse(row.Cells["Dias de Preparação"].Value.ToString(), out horasValue);

                    if (isHorasNumeric)
                    {
                        row.Cells["Dias de Preparação"].Value = $"{horasValue} dias";
                    }
                }
            }

        }
        private string SubstituirSeparadorDecimal(string valor)
        {
            return valor.Replace(",", ".");
        }
        private void RemoverSufixosNasColunasOracamentacao()
        {
            foreach (DataGridViewRow row in DataGridViewOrcamentacaoObras.Rows)
            {
                if (row.Cells["KG Estrutura"].Value != DBNull.Value && row.Cells["KG Estrutura"].Value != null)
                {
                    string kgValueStr = row.Cells["KG Estrutura"].Value.ToString();
                    kgValueStr = kgValueStr.Replace(" kg", "").Trim();

                    double kgValue;
                    bool isKgNumeric = Double.TryParse(kgValueStr, out kgValue);

                    if (isKgNumeric)
                    {
                        row.Cells["KG Estrutura"].Value = kgValue;
                    }
                }

                if (row.Cells["Horas Estrutura"].Value != DBNull.Value && row.Cells["Horas Estrutura"].Value != null)
                {
                    string horasValueStr = row.Cells["Horas Estrutura"].Value.ToString();
                    horasValueStr = horasValueStr.Replace(" h", "").Trim();

                    double horasValue;
                    bool isHorasNumeric = Double.TryParse(horasValueStr, out horasValue);

                    if (isHorasNumeric)
                    {
                        row.Cells["Horas Estrutura"].Value = horasValue;
                    }
                }

                if (row.Cells["Valor Estrutura"].Value != DBNull.Value && row.Cells["Valor Estrutura"].Value != null)
                {
                    string valorStr = row.Cells["Valor Estrutura"].Value.ToString();
                    valorStr = valorStr.Replace(" €", "").Replace('.', ',').Trim();

                    double valorValue;
                    bool isValorNumeric = Double.TryParse(valorStr, out valorValue);

                    if (isValorNumeric)
                    {
                        row.Cells["Valor Estrutura"].Value = valorValue;
                    }
                }

                if (row.Cells["KG/Euro Estrutura"].Value != DBNull.Value && row.Cells["KG/Euro Estrutura"].Value != null)
                {
                    string valorStr = row.Cells["KG/Euro Estrutura"].Value.ToString();
                    valorStr = valorStr.Replace(" €", "").Replace('.', ',').Trim();

                    double valorValue;
                    bool isValorNumeric = Double.TryParse(valorStr, out valorValue);

                    if (isValorNumeric)
                    {
                        row.Cells["KG/Euro Estrutura"].Value = valorValue;
                    }
                }

                if (row.Cells["Horas Revestimentos"].Value != DBNull.Value && row.Cells["Horas Revestimentos"].Value != null)
                {
                    string horasValueStr = row.Cells["Horas Revestimentos"].Value.ToString();
                    horasValueStr = horasValueStr.Replace(" h", "").Trim();

                    double horasValue;
                    bool isHorasNumeric = Double.TryParse(horasValueStr, out horasValue);

                    if (isHorasNumeric)
                    {
                        row.Cells["Horas Revestimentos"].Value = horasValue;
                    }
                }

                if (row.Cells["Valor Revestimentos"].Value != DBNull.Value && row.Cells["Valor Revestimentos"].Value != null)
                {
                    string valorStr = row.Cells["Valor Revestimentos"].Value.ToString();
                    valorStr = valorStr.Replace(" €", "").Replace('.', ',').Trim();

                    double valorValue;
                    bool isValorNumeric = Double.TryParse(valorStr, out valorValue);

                    if (isValorNumeric)
                    {
                        row.Cells["Valor Revestimentos"].Value = valorValue;
                    }
                }

                if (row.Cells["Total Horas"].Value != DBNull.Value && row.Cells["Total Horas"].Value != null)
                {
                    string horasValueStr = row.Cells["Total Horas"].Value.ToString();
                    horasValueStr = horasValueStr.Replace(" h", "").Trim();

                    double horasValue;
                    bool isHorasNumeric = Double.TryParse(horasValueStr, out horasValue);

                    if (isHorasNumeric)
                    {
                        row.Cells["Total Horas"].Value = horasValue;
                    }
                }

                if (row.Cells["Total Valor"].Value != DBNull.Value && row.Cells["Total Valor"].Value != null)
                {
                    string valorStr = row.Cells["Total Valor"].Value.ToString();
                    valorStr = valorStr.Replace(" €", "").Replace('.', ',').Trim();

                    double valorValue;
                    bool isValorNumeric = Double.TryParse(valorStr, out valorValue);

                    if (isValorNumeric)
                    {
                        row.Cells["Total Valor"].Value = valorValue;
                    }
                }
            }
        }
        private void VerificarEAtualizarOuSalvar()
        {
            if (DataGridViewOrcamentacaoObras.SelectedRows.Count > 0)
            {
                string idSelecionado = DataGridViewOrcamentacaoObras.SelectedRows[0].Cells["ID"].Value.ToString();

                if (!string.IsNullOrEmpty(idSelecionado))
                {
                    AtualizarOrcamentoNaBD();
                }
                else
                {
                    SalvarOrcamentacaoBD();
                }
            }
            else
            {
                MessageBox.Show("Selecione uma linha.");
            }
        }
        private void Modificartabelas()
        {
            int alturaCabecalho = 45;
            int alturaLinha = 28;
            int numeroLinhas = DataGridViewOrcamentacaoObras.Rows.Count;
            DataGridViewOrcamentacaoObrasTotal.BringToFront();
            DataGridViewRealObrasTotal.BringToFront();
            DataGridViewConclusaoObrasTotal.BringToFront();
            int alturaTotal = alturaCabecalho + (numeroLinhas * alturaLinha);
            alturaTotal = Math.Min(alturaTotal, 2100);
            PanelTabelas.Size = new Size(2342, alturaTotal + 35);
            DataGridViewOrcamentacaoObras.Size = new Size(1150, alturaTotal);
            DataGridViewRealObras.Size = new Size(840, alturaTotal + 20);
            DataGridViewConclusaoObras.Size = new Size(340, alturaTotal);    
        }
        private bool isSynchronizingSelection = false; private void DataGridViewRealObras_SelectionChanged(object sender, EventArgs e)
        {
            if (isSynchronizingSelection) return;

            if (DataGridViewRealObras.SelectedRows.Count > 0)
            {
                int selectedIndex = DataGridViewRealObras.SelectedRows[0].Index;

                isSynchronizingSelection = true;

                if (DataGridViewOrcamentacaoObras.Rows.Count > selectedIndex)
                {
                    DataGridViewOrcamentacaoObras.ClearSelection();
                    DataGridViewOrcamentacaoObras.Rows[selectedIndex].Selected = true;
                }

                if (DataGridViewConclusaoObras.Rows.Count > selectedIndex)
                {
                    DataGridViewConclusaoObras.ClearSelection();
                    DataGridViewConclusaoObras.Rows[selectedIndex].Selected = true;
                }

                isSynchronizingSelection = false;
            }
        }
        private void DataGridViewRealObrasNome_SelectionChanged(object sender, EventArgs e)
        {
            if (isSynchronizingSelection) return;

            if (DataGridViewOrcamentacaoObras.SelectedRows.Count > 0)
            {
                int selectedIndex = DataGridViewOrcamentacaoObras.SelectedRows[0].Index;

                isSynchronizingSelection = true;

                if (DataGridViewRealObras.Rows.Count > selectedIndex)
                {
                    DataGridViewRealObras.ClearSelection();
                    DataGridViewRealObras.Rows[selectedIndex].Selected = true;
                }

                if (DataGridViewConclusaoObras.Rows.Count > selectedIndex)
                {
                    DataGridViewConclusaoObras.ClearSelection();
                    DataGridViewConclusaoObras.Rows[selectedIndex].Selected = true;
                }

                isSynchronizingSelection = false;
            }
        }
        private bool isSynchronizingScroll = false; private void Guna2VScrollBar_Scroll(object sender, ScrollEventArgs e)
        {
            int newValue = e.NewValue;

            if (DataGridViewRealObras.FirstDisplayedScrollingRowIndex != newValue)
            {
                DataGridViewRealObras.FirstDisplayedScrollingRowIndex = newValue;
            }

            if (DataGridViewOrcamentacaoObras.FirstDisplayedScrollingRowIndex != newValue)
            {
                DataGridViewOrcamentacaoObras.FirstDisplayedScrollingRowIndex = newValue;
            }

            if (DataGridViewConclusaoObras.FirstDisplayedScrollingRowIndex != newValue)
            {
                DataGridViewConclusaoObras.FirstDisplayedScrollingRowIndex = newValue;
            }
        }
        private void DataGridViewRealObras_Scroll(object sender, ScrollEventArgs e)
        {
            if (isSynchronizingScroll) return;

            isSynchronizingScroll = true;

            guna2VScrollBar1.Value = DataGridViewRealObras.FirstDisplayedScrollingRowIndex;

            isSynchronizingScroll = false;
        }
        private void DataGridViewOrcamentacaoObras_Scroll(object sender, ScrollEventArgs e)
        {
            if (isSynchronizingScroll) return;

            isSynchronizingScroll = true;

            guna2VScrollBar1.Value = DataGridViewOrcamentacaoObras.FirstDisplayedScrollingRowIndex;

            isSynchronizingScroll = false;
        }
        private void DataGridViewConclusaoObras_Scroll(object sender, ScrollEventArgs e)
        {
            if (isSynchronizingScroll) return;

            isSynchronizingScroll = true;

            guna2VScrollBar1.Value = DataGridViewConclusaoObras.FirstDisplayedScrollingRowIndex;

            isSynchronizingScroll = false;
        }
        private void DataGridViewRealObras_ScrollHorizontal(object sender, ScrollEventArgs e)
        {
            if (isSynchronizingScroll) return;
            isSynchronizingScroll = true;
            isSynchronizingScroll = false;
        }
        private void DataGridViewRealObrasTotal_ScrollHorizontal(object sender, ScrollEventArgs e)
        {
            if (isSynchronizingScroll) return;
            isSynchronizingScroll = true;
            isSynchronizingScroll = false;
        }
        private void InicializarSincronizacaoDeRolagem()
        {
            DataGridViewRealObras.Scroll += DataGridViewRealObras_Scroll;
            DataGridViewOrcamentacaoObras.Scroll += DataGridViewOrcamentacaoObras_Scroll;
            DataGridViewConclusaoObras.Scroll += DataGridViewConclusaoObras_Scroll;
            DataGridViewRealObras.Scroll += DataGridViewRealObras_ScrollHorizontal;
            DataGridViewRealObrasTotal.Scroll += DataGridViewRealObrasTotal_ScrollHorizontal;
            guna2VScrollBar1.Scroll += Guna2VScrollBar_Scroll;
        }
        private void DataGridViewOrcamentacaoObras_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            if (e.ColumnIndex == DataGridViewOrcamentacaoObras.Columns["Tipologia"]?.Index)
            {
                DataGridViewComboBoxCell comboBoxCell = new DataGridViewComboBoxCell();
                comboBoxCell.Items.Add("Porticado");
                comboBoxCell.Items.Add("Treliçado");
                comboBoxCell.Items.Add("Revestimentos");

                var currentValue = DataGridViewOrcamentacaoObras[e.ColumnIndex, e.RowIndex].Value;
                if (currentValue != null && comboBoxCell.Items.Contains(currentValue))
                {
                    comboBoxCell.Value = currentValue;
                }
                else
                {
                    comboBoxCell.Value = "Porticado";
                }

                DataGridViewOrcamentacaoObras[e.ColumnIndex, e.RowIndex] = comboBoxCell;
            }
            if (e.ColumnIndex == DataGridViewOrcamentacaoObras.Columns["Preparador Responsavel"]?.Index)
            {
                DataGridViewComboBoxCell comboBoxCell = new DataGridViewComboBoxCell();
                foreach (var item in ComboBoxPreparadorAdd.Items)
                {
                    comboBoxCell.Items.Add(item);
                }
                var currentValue = DataGridViewOrcamentacaoObras[e.ColumnIndex, e.RowIndex].Value;
                if (currentValue != null && comboBoxCell.Items.Contains(currentValue))
                {
                    comboBoxCell.Value = currentValue;
                }
                else
                {
                    comboBoxCell.Value = ComboBoxPreparadorAdd.Items[0];
                }
                DataGridViewOrcamentacaoObras[e.ColumnIndex, e.RowIndex] = comboBoxCell;
            }
        }
        private void DataGridViewOrcamentacaoObras_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == DataGridViewOrcamentacaoObras.Columns["Tipologia"]?.Index)
            {
                var selectedValue = DataGridViewOrcamentacaoObras[e.ColumnIndex, e.RowIndex].Value.ToString();
                DataGridViewTextBoxCell textCell = new DataGridViewTextBoxCell();
                textCell.Value = selectedValue;
                DataGridViewOrcamentacaoObras[e.ColumnIndex, e.RowIndex] = textCell;
            }
            if (e.ColumnIndex == DataGridViewOrcamentacaoObras.Columns["Preparador Responsavel"]?.Index)
            {
                var selectedValue = DataGridViewOrcamentacaoObras[e.ColumnIndex, e.RowIndex].Value.ToString();
                DataGridViewTextBoxCell textCell = new DataGridViewTextBoxCell();
                textCell.Value = selectedValue;
                DataGridViewOrcamentacaoObras[e.ColumnIndex, e.RowIndex] = textCell;
            }
        }
        public DataTable DataGridViewToDataTable(DataGridView dataGridView, bool incluirCabecalho)
        {
            DataTable dataTable = new DataTable();

            if (dataGridView.Columns.Count == 0)
            {
                throw new InvalidOperationException("O DataGridView não tem colunas.");
            }
            if (incluirCabecalho)
            {
                foreach (DataGridViewColumn column in dataGridView.Columns)
                {
                    if (column.Visible)
                    {
                        dataTable.Columns.Add(column.HeaderText);
                    }
                }
            }
            foreach (DataGridViewRow row in dataGridView.Rows)
            {
                if (!row.IsNewRow)
                {
                    DataRow dataRow = dataTable.NewRow();
                    int cellIndex = 0;

                    foreach (DataGridViewCell cell in row.Cells)
                    {
                        if (cell.OwningColumn.Visible)
                        {
                            if (cellIndex < dataTable.Columns.Count)
                            {
                                dataRow[cellIndex] = cell.Value;
                            }
                            cellIndex++;
                        }
                    }

                    dataTable.Rows.Add(dataRow);
                }
            }
            return dataTable;
        }
        public void ExportarTabelaValoresExcel()
        {
            string filePath = @"C:\r\RegistrosCompletos.xlsx";
            DataGridView dataGridViewOrcamentacaoObras = this.DataGridViewOrcamentacaoObras;
            DataGridView dataGridViewRealObras = this.DataGridViewRealObras;
            DataGridView dataGridViewConclusaoObras = this.DataGridViewConclusaoObras;

            DataGridView DataGridViewOrcamentacaoObrasTotal = this.DataGridViewOrcamentacaoObrasTotal;
            DataGridView DataGridViewRealObras = this.DataGridViewRealObrasTotal;
            DataGridView DataGridViewConclusaoObrasTotal = this.DataGridViewConclusaoObrasTotal;

            ExcelExport excelExport = new ExcelExport();
            DataTable dataTableOrcamentacaoObras = DataGridViewToDataTable(dataGridViewOrcamentacaoObras, true); // Cabeçalho incluído
            DataTable dataTableRealObras = DataGridViewToDataTable(dataGridViewRealObras, true);
            DataTable dataTableConclusaoObras = DataGridViewToDataTable(dataGridViewConclusaoObras, true);

            DataTable dataTableOrcamentacaoObrasTotal = DataGridViewToDataTable(DataGridViewOrcamentacaoObrasTotal, false); // Sem cabeçalho
            DataTable dataTableRealObrasTotal = DataGridViewToDataTable(DataGridViewRealObras, false);
            DataTable dataTableConclusaoObrasTotal = DataGridViewToDataTable(DataGridViewConclusaoObrasTotal, false);

            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("RegistrosCompletos");

                excelExport.ExportarParaExcelTabela(dataTableOrcamentacaoObras, worksheet, 2);

                excelExport.ExportarParaExcelTabela(dataTableRealObras, worksheet, 15);

                excelExport.ExportarParaExcelTabela(dataTableConclusaoObras, worksheet, 44);

                Calculartotais(filePath);

                FileInfo fi = new FileInfo(filePath);
                package.SaveAs(fi);
            }
            ExportarParaExcelComTotais(filePath, DataGridViewOrcamentacaoObrasTotal, DataGridViewRealObras, DataGridViewConclusaoObrasTotal);

            try
            {
                System.Diagnostics.Process.Start(filePath);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erro ao tentar abrir o arquivo Excel: " + ex.Message, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        public void ExportarParaExcelComTotais(string filePath, DataGridView DataGridViewOrcamentacaoObrasTotal, DataGridView DataGridViewRealObras, DataGridView DataGridViewConclusaoObrasTotal)
        {
            if (!File.Exists(filePath))
            {
                MessageBox.Show("Arquivo Excel não encontrado.");
                return;
            }
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets[0];
                int lastRow = worksheet.Dimension.End.Row + 1;
                ExportarTotaisParaExcel(worksheet, DataGridViewOrcamentacaoObrasTotal, lastRow, 7, 14);
                ExportarTotaisParaExcel(worksheet, DataGridViewRealObras, lastRow, 14, 43);
                ExportarTotaisParaExcel(worksheet, DataGridViewConclusaoObrasTotal, lastRow, 43, 47);
                package.Save();
            }
            MessageBox.Show("Totais exportados com sucesso!");
        }
        private void ExportarTotaisParaExcel(ExcelWorksheet worksheet, DataGridView dataGridView, int startRow, int startColumn, int endColumn)
        {
            for (int col = startColumn; col <= endColumn; col++)
            {
                double total = 0;
                bool hasNumericData = false;

                for (int row = 0; row < dataGridView.Rows.Count; row++)
                {
                    if (col - 1 < dataGridView.Rows[row].Cells.Count)
                    {
                        var cellValue = dataGridView.Rows[row].Cells[col - 1].Value;
                        if (cellValue != null && double.TryParse(cellValue.ToString(), out double value)) // Ajusta índice para o Excel
                        {
                            total += value;
                            hasNumericData = true;
                        }
                    }
                }
                if (hasNumericData)
                {
                    var cell = worksheet.Cells[startRow, col];
                    cell.Value = total;
                    cell.Style.Font.Bold = true;
                    cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    cell.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(217, 217, 217)); // Cor de fundo
                    cell.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                    cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    cell.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                }
            }
        }
        private void ExportExcelRegistodeTodos()
        {
            string query = @"
                            SELECT [Numero da Obra], Preparador, [Data da Tarefa], [Qtd de Hora], [Hora Inicial], [Hora Final], Prioridade, Tarefa
                            FROM dbo.RegistoTempo
                            ORDER BY [Numero da Obra]";

            ComunicaBD comunicaBD = new ComunicaBD();
            ExcelExport excelExport = new ExcelExport();
            SqlCommand command = new SqlCommand(query, comunicaBD.GetConnection());
            comunicaBD.ConectarBD();
            DataTable dataTable = comunicaBD.BuscarRegistros(command);
            string filePath = $@"C:\r\Registros.xlsx";
            excelExport.ExportarParaExcelTodos(dataTable, filePath);
            comunicaBD.DesonectarBD();
            try
            {
                System.Diagnostics.Process.Start(filePath);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erro ao tentar abrir o arquivo Excel: " + ex.Message, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void Calculartotais(string filePath)
        {
            try
            {
                if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
                {
                    MessageBox.Show("Caminho do arquivo inválido!", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                FileInfo fileInfo = new FileInfo(filePath);
                using (var package = new ExcelPackage(fileInfo))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                    int ultimaLinha = worksheet.Dimension.End.Row;
                    double total = 0;
                    int colunaValor = 1;
                    for (int linha = 1; linha <= ultimaLinha; linha++)
                    {
                        string valorCelula = worksheet.Cells[linha, colunaValor].Text;
                        string valorLimpo = LimparValor(valorCelula);
                        if (double.TryParse(valorLimpo, out double valorNum))
                        {
                            total += valorNum;
                        }
                    }
                    worksheet.Cells[ultimaLinha + 1, colunaValor].Value = total;
                    package.Save();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erro ao processar o arquivo: " + ex.Message, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private string LimparValor(string valor)
        {
            return System.Text.RegularExpressions.Regex.Replace(valor, @"[^\d.,]", "");
        }
        private double ProcessarValor(object value)
        {
            if (value == null || value == DBNull.Value)
            {
                return 0;
            }
            string strValue = value.ToString().Trim();
            if (string.IsNullOrEmpty(strValue))
            {
                return 0;
            }
            strValue = strValue.Replace(",", ".");
            double result = 0;
            if (double.TryParse(strValue, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out result))
            {
                return result;
            }
            else
            {
                return 0;
            }
        }
        private double CalcularPercentagem(string valorPercentagem)
        {
            if (string.IsNullOrEmpty(valorPercentagem))
            {
                return 0;
            }
            valorPercentagem = valorPercentagem.Replace("%", "").Trim();
            valorPercentagem = valorPercentagem.Replace(",", ".");
            double resultado;
            if (double.TryParse(valorPercentagem, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out resultado))
            {
                return resultado;
            }
            else
            {
                return 0;
            }
        }
        private void PontoporvirgulaOrcamentacao()
        {
            foreach (DataGridViewRow row in DataGridViewOrcamentacaoObras.Rows)
            {
                foreach (DataGridViewCell cell in row.Cells)
                {
                    if (cell.Value != null && cell.Value is string)
                    {
                        string cellValue = cell.Value.ToString();

                        cell.Value = cellValue.Replace(",", ".");
                    }
                }
            }
        }
        private void ComboBoxTipologiaFiltro_SelectedIndexChanged(object sender, EventArgs e)
        {            
            FiltrarTipologia();
            VisualizarGraficoTotalHoras();
            VisualizarGraficoObrasHorasTipologia();
            VisualizarGraficoObrasValorTipologia();
            VisualizarGraficoObrasPercentagemTipologia();
            VisualizarGraficoHorasTipologiaeAno();
            VisualizarGraficoPiePercentagemTipologiaeAno();
            Modificartabelas();
            Calcular();
            LimparColunaUltimaLinha(DataGridViewOrcamentacaoObras);
            AdicionarSufixosNasColunas();
        }
        private void ComboBoxAnoAdd_SelectedIndexChanged(object sender, EventArgs e)
        {            
            FiltrarPorAno();
            VisualizarGraficoHorasTipologiaeAno();
            VisualizarGraficoPiePercentagemTipologiaeAno();
            VisualizarGraficoObrasHorasAno();
            VisualizarGraficoObrasValorAno();
            VisualizarGraficoObrasPercentagemAno();
            Modificartabelas();
            Calcular();
            LimparColunaUltimaLinha(DataGridViewOrcamentacaoObras);
            AdicionarSufixosNasColunas();
        }
        private void ButtonLimaprtaabela_Click(object sender, EventArgs e)
        {
            pictureBoxGif.Visible = true;
            pictureBoxGif.BringToFront();

            Timer timer = new Timer();
            timer.Interval = 1000;
            timer.Tick += (s, ev) =>
            {
                timer.Stop();
                pictureBoxGif.Visible = false;
            };
            timer.Start();           
            ComunicarTabelas();
            CarregarGraficos();
            Modificartabelas();
            Calcular();
            LimparColunaUltimaLinha(DataGridViewOrcamentacaoObras);
            AdicionarSufixosNasColunas();
        }
        private void ButtonInserirAno_Click(object sender, EventArgs e)
        {
            InserirAnofecho();
            ComunicarTabelas();
            Calcular();
            LimparColunaUltimaLinha(DataGridViewOrcamentacaoObras);
            AdicionarSufixosNasColunas();
        }
        private void ButtonInserirTipologia_Click(object sender, EventArgs e)
        {
            InserirTipologia();
            ComunicarTabelas();
            Calcular();
            LimparColunaUltimaLinha(DataGridViewOrcamentacaoObras);
            AdicionarSufixosNasColunas();
        }
        private void Buttonlimparanofecho_Click(object sender, EventArgs e)
        {
            pictureBoxGif1.Visible = true;
            pictureBoxGif1.BringToFront();

            Timer timer = new Timer();
            timer.Interval = 1000;
            timer.Tick += (s, ev) =>
            {
                timer.Stop();
                pictureBoxGif1.Visible = false;
            };
            timer.Start();           
            LimparAnoFecho();
            CarregarGraficos();
            Modificartabelas();
            Calcular();
            LimparColunaUltimaLinha(DataGridViewOrcamentacaoObras);
            AdicionarSufixosNasColunas();
        }
        private void ButtonAtualizarKG_Click(object sender, EventArgs e)
        {
            AtualizarTabelaRealnaBdManual();
            AtualizarDados();
            CalcularTabelas();
        }
        private void ButtonPrecohora_Click(object sender, EventArgs e)
        {
            string novoPreco = TextBoxprecoHora.Text;
            if (!string.IsNullOrWhiteSpace(novoPreco))
            {
                LabelPrecoOrcamentado.Text = novoPreco;
                Properties.Settings.Default.PrecoOrcamentado = novoPreco;
                Properties.Settings.Default.Save();
            }
            else
            {
                MessageBox.Show("Digite um valor no campo Preço/Hora.");
            }
        }
        private void ButtonAtualizarbd_Click(object sender, EventArgs e)
        {
            RemoverSufixosNasColunasOracamentacao();
            VerificarEAtualizarOuSalvar();
            VisualizarTabelaOrcamentacao();
            VisualizarTabelaReal();
            VisualizarTabelaConcluido();
            AdicionarSufixosNasColunas();
            Modificartabelas();
            Calcular();
            LimparColunaUltimaLinha(DataGridViewOrcamentacaoObras);
            AdicionarSufixosNasColunas();
        }
        private void ButtonExportExcelTodas_Click_1(object sender, EventArgs e)
        {
            ExportExcelRegistodeTodos();
        }
        private void ButtonExportExcelTabelas_Click(object sender, EventArgs e)
        {
            ExportarTabelaValoresExcel();
        }
        private void Calcular()
        {
            CalcularTotaisorcamentacao();
            CalcularTotaisReais();
            CalcularTotaisConclus();
        }        
        private void CalcularTotaisorcamentacao()
        {
            if (DataGridViewOrcamentacaoObras.Rows.Count == 0)
                return;

            int ultimaLinha = DataGridViewOrcamentacaoObras.Rows.Count - 1;
            DataGridViewRow totalRow = DataGridViewOrcamentacaoObras.Rows[ultimaLinha];

            int totalColunas = DataGridViewOrcamentacaoObras.Columns.Count;

            for (int coluna = 0; coluna < totalColunas; coluna++)
            {
                decimal soma = 0;

                foreach (DataGridViewRow row in DataGridViewOrcamentacaoObras.Rows)
                {
                    if (row.IsNewRow || row.Index == ultimaLinha)
                        continue;

                    object valor = row.Cells[coluna].Value;
                    if (valor != null)
                    {
                        string texto = valor.ToString().Trim().Replace(".", ",");
                        if (decimal.TryParse(texto, NumberStyles.Any, CultureInfo.CurrentCulture, out decimal num))
                        {
                            soma += num;
                        }
                    }
                }

                totalRow.Cells[coluna].Value = soma;
            }
                       
            totalRow.DefaultCellStyle.BackColor = Color.GhostWhite;
            totalRow.DefaultCellStyle.ForeColor = Color.Black;
            totalRow.DefaultCellStyle.Font = new Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold);
        }
        private void CalcularTotaisReais()
        {
            if (DataGridViewRealObras.Rows.Count == 0)
                DataGridViewRealObras.Rows.Add();
            int ultimaLinha = DataGridViewRealObras.Rows.Count - 1;
            DataGridViewRow totalRow = DataGridViewRealObras.Rows[ultimaLinha];
            for (int colunaOrigem = 0; colunaOrigem < DataGridViewRealObras.Columns.Count; colunaOrigem++)
            {
                decimal soma = 0;
                foreach (DataGridViewRow row in DataGridViewRealObras.Rows)
                {
                    if (row.IsNewRow || row.Index == ultimaLinha) continue;
                    object valor = row.Cells[colunaOrigem].Value;
                    if (valor != null)
                    {
                        string texto = valor.ToString().Trim().Replace("%", "").Replace(".", ",");
                        if (decimal.TryParse(texto, NumberStyles.Any, CultureInfo.CurrentCulture, out decimal num))
                        {
                            soma += num;
                        }
                    }
                }
                totalRow.Cells[colunaOrigem].Value = soma;
            }
            CalcularMediaColuna(DataGridViewRealObras, 7);
            int[] colunasParaMedia = { 8, 11, 14, 17, 20, 23, 26, 29 };
            foreach (int coluna in colunasParaMedia)
            {
                CalcularMediaColuna(DataGridViewRealObras, coluna);
                int ultimaLinhamedia = DataGridViewRealObras.Rows.Count - 1;
                var totalRowmedia = DataGridViewRealObras.Rows[ultimaLinha];
                if (totalRowmedia.Cells[coluna].Value != null)
                    totalRowmedia.Cells[coluna].Value = totalRow.Cells[coluna].Value.ToString() + "%";            
            }
            totalRow.DefaultCellStyle.BackColor = Color.GhostWhite;
            totalRow.DefaultCellStyle.ForeColor = Color.Black;
            totalRow.DefaultCellStyle.Font = new Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold);
        }
        private void CalcularTotaisConclus()
        {
            if (DataGridViewConclusaoObras.Rows.Count == 0)
                DataGridViewConclusaoObras.Rows.Add();
            int ultimaLinha = DataGridViewConclusaoObras.Rows.Count - 1;
            DataGridViewRow totalRow = DataGridViewConclusaoObras.Rows[ultimaLinha];

            for (int colunaOrigem = 0; colunaOrigem < DataGridViewConclusaoObras.Columns.Count; colunaOrigem++)
            {
                decimal soma = 0;
                foreach (DataGridViewRow row in DataGridViewConclusaoObras.Rows)
                {
                    if (row.IsNewRow || row.Index == ultimaLinha) continue;

                    object valor = row.Cells[colunaOrigem].Value;
                    if (valor != null)
                    {
                        string texto = valor.ToString().Trim().Replace("%", "").Replace(".", ",");
                        if (decimal.TryParse(texto, NumberStyles.Any, CultureInfo.CurrentCulture, out decimal num))
                        {
                            soma += num;
                        }
                    }
                }
                totalRow.Cells[colunaOrigem].Value = soma;
            }
            int[] colunasParaMedia = { 6 };
            foreach (int coluna in colunasParaMedia)
            {
                CalcularMediaColuna(DataGridViewConclusaoObras, coluna);
                int ultimaLinhamedia = DataGridViewConclusaoObras.Rows.Count - 1;
                var totalRowmedia = DataGridViewConclusaoObras.Rows[ultimaLinha];
                if (totalRowmedia.Cells[coluna].Value != null)
                    totalRowmedia.Cells[coluna].Value = totalRow.Cells[coluna].Value.ToString() + "%";
            }
            totalRow.DefaultCellStyle.BackColor = Color.GhostWhite;
            totalRow.DefaultCellStyle.ForeColor = Color.Black;
            totalRow.DefaultCellStyle.Font = new Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold);
        }
        private void CalcularMediaColuna(DataGridView dgv, int indiceColuna)
        {
            if (dgv.Rows.Count == 0)
                return;
            int ultimaLinha = dgv.Rows.Count - 1;
            DataGridViewRow totalRow = dgv.Rows[ultimaLinha];
            object valor = totalRow.Cells[indiceColuna].Value;
            if (valor != null)
            {
                string texto = valor.ToString().Trim().Replace(".", ",");
                if (decimal.TryParse(texto, NumberStyles.Any, CultureInfo.CurrentCulture, out decimal num))
                {
                    int linhasValidas = dgv.Rows.Count - 1;
                    if (linhasValidas > 0)
                    {
                        decimal resultado = num / linhasValidas;
                        resultado = Math.Round(resultado, 1, MidpointRounding.AwayFromZero);
                        totalRow.Cells[indiceColuna].Value = resultado;
                    }
                }
            }
        }
        private void ConfigurarColunasOrcamentacao()
        {
            if (DataGridViewOrcamentacaoObras.Columns.Count == 0) return;
            DataGridViewOrcamentacaoObras.Columns["Id"].Visible = false;
            DataGridViewOrcamentacaoObras.Columns["Ano de fecho"].Width = 40;
            DataGridViewOrcamentacaoObras.Columns["Numero da Obra"].Width = 60;
            DataGridViewOrcamentacaoObras.Columns["Nome da Obra"].Width = 120;
            DataGridViewOrcamentacaoObras.Columns["Preparador Responsavel"].Width = 70;
            DataGridViewOrcamentacaoObras.Columns["Tipologia"].Width = 65;
            DataGridViewOrcamentacaoObras.Columns["KG Estrutura"].Width = 70;
            DataGridViewOrcamentacaoObras.Columns["Horas Estrutura"].Width = 50;
            DataGridViewOrcamentacaoObras.Columns["Valor Estrutura"].Width = 75;
            DataGridViewOrcamentacaoObras.Columns["KG/Euro Estrutura"].Width = 40;
            DataGridViewOrcamentacaoObras.Columns["Horas Revestimentos"].Width = 50;
            DataGridViewOrcamentacaoObras.Columns["Valor Revestimentos"].Width = 50;
            DataGridViewOrcamentacaoObras.Columns["Total Horas"].Width = 50;
            DataGridViewOrcamentacaoObras.Columns["Total Valor"].Width = 90;
            int index = 0;
            DataGridViewOrcamentacaoObras.Columns["Ano de fecho"].DisplayIndex = index++;
            DataGridViewOrcamentacaoObras.Columns["Numero da Obra"].DisplayIndex = index++;
            DataGridViewOrcamentacaoObras.Columns["Nome da Obra"].DisplayIndex = index++;
            DataGridViewOrcamentacaoObras.Columns["Preparador Responsavel"].DisplayIndex = index++;
            DataGridViewOrcamentacaoObras.Columns["Tipologia"].DisplayIndex = index++;
            DataGridViewOrcamentacaoObras.Columns["KG Estrutura"].DisplayIndex = index++;
            DataGridViewOrcamentacaoObras.Columns["Horas Estrutura"].DisplayIndex = index++;
            DataGridViewOrcamentacaoObras.Columns["Valor Estrutura"].DisplayIndex = index++;
            DataGridViewOrcamentacaoObras.Columns["KG/Euro Estrutura"].DisplayIndex = index++;
            DataGridViewOrcamentacaoObras.Columns["Horas Revestimentos"].DisplayIndex = index++;
            DataGridViewOrcamentacaoObras.Columns["Valor Revestimentos"].DisplayIndex = index++;
            DataGridViewOrcamentacaoObras.Columns["Total Horas"].DisplayIndex = index++;
            DataGridViewOrcamentacaoObras.Columns["Total Valor"].DisplayIndex = index++;

            DataGridViewOrcamentacaoObras.Columns["KG/Euro Estrutura"].HeaderText = "Euro/kg Estrutura";

        }
        private void Limpar(DataGridView dgv)
        {
            if (dgv.Rows.Count == 0)
                return; 
            int ultimaLinhaIndex = dgv.Rows.Count - 1;
            DataGridViewRow ultimaLinha = dgv.Rows[ultimaLinhaIndex];
            for (int i = 0; i <= 5 && i < ultimaLinha.Cells.Count; i++)
            {
                ultimaLinha.Cells[i].Value = null;
            }
        }
        private void LimparColunaUltimaLinha(DataGridView dgv)
        {
            if (dgv.Rows.Count == 0)
                return;
            int ultimaLinhaIndex = dgv.Rows.Count - 1;
            DataGridViewRow ultimaLinha = dgv.Rows[ultimaLinhaIndex];
            for (int i = 0; i <= 3 && i < ultimaLinha.Cells.Count; i++)
            {
                ultimaLinha.Cells[i].Value = null;
            }
            int[] colunasParaLimpar = { 12, 13 };
            foreach (int col in colunasParaLimpar)
            {
                if (col < ultimaLinha.Cells.Count)
                    ultimaLinha.Cells[col].Value = null;
            }
        }
        private void checkanofecho_Click(object sender, EventArgs e)
        {
            if (checkanofecho.Checked)
            {
                labelanofecho.Text = "Com Ano de fecho";
            }
            else
            {
                labelanofecho.Text = "Sem Ano de fecho";
            }
        }
    }
}





