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

namespace OfelizCM
{
    public partial class Frm_TodasObras : Form
    {
        public Frm_TodasObras()
        {
            InitializeComponent();
        }

        private void Frm_TodasObras_Load(object sender, EventArgs e)
        {
            this.totalObrasTableAdapter.Fill(this.tempoPreparacaoDataSet.TotalObras);
            ComunicarTabelas();
            AtualizarDados();
            CalcularTabelas();
            CarregarGraficos();
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
            DataGridViewRealObras.ScrollBars = ScrollBars.None;
            VerificarUsuario();
            InicializarSincronizacaoDeRolagem();
            CarregarPastasNaComboBoxAno();
            FormatacaodosGraficos();
            AdicionarSufixosNasColunas();
            CarregarTipologiaNaComboBox(); 

            chartObrasHoras.ChartAreas[0].AxisX.IsMarginVisible = true;
            chartObrasHoras.ChartAreas[0].AxisX.LabelStyle.Angle = -45;  // Gira os rótulos do eixo X em -45 graus
            chartObrasHoras.ChartAreas[0].AxisX.Interval = 1;  // Exibe um rótulo a cada 1 valor no eixo X                      

            chartTotalPercentagem.ChartAreas[0].AxisX.IsMarginVisible = true;
            chartTotalPercentagem.ChartAreas[0].AxisX.LabelStyle.Angle = -45;  
            chartTotalPercentagem.ChartAreas[0].AxisX.Interval = 1;  
        }

        private void ComunicarTabelas()
        {
            ComunicaBDparaTabelaOrcamentacao();
            ComunicaBDparaTabelaReal();
            ComunicaBDparaTabelaConcluido();
            ComunicaBDparaTabelaOrcamentacaoTotal();
            ComunicaBDparaTabelaRealTotais();
            ComunicaBDparaTabelaConclusaoTotal();
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
            CarregarGraficoObras();
            CarregarGraficoObras2();
            CarregarGraficoObrasvalor();
            CarregarGraficoObrasPercentagem();
            CarregarGraficoPiePercentagem();
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
                            guna2ImageButton7.Visible = true;
                            ButtonExportExcelTodas.Visible = true;
                            guna2ImageButton5.Visible = true;
                            guna2ImageButton4.Visible = true;
                            guna2ImageButton3.Visible = true;
                            guna2ImageButton6.Visible = true;
                            guna2ImageButton2.Visible = true;

                        }
                        else
                        {
                            guna2ImageButton7.Visible = false;
                            ButtonExportExcelTodas.Visible = false;
                            guna2ImageButton5.Visible = false;
                            guna2ImageButton4.Visible = false;
                            guna2ImageButton3.Visible = false;
                            guna2ImageButton6.Visible = false;
                            guna2ImageButton2.Visible = false;

                        }
                    }
                    else
                    {

                        guna2ImageButton7.Visible = false;
                        ButtonExportExcelTodas.Visible = false;
                        guna2ImageButton5.Visible = false;
                        guna2ImageButton4.Visible = false;
                        guna2ImageButton3.Visible = false;
                        guna2ImageButton6.Visible = false;
                        guna2ImageButton2.Visible = false;
                    }
                    string nomeUsuario2 = Properties.Settings.Default.NomeUsuario;

                    if (nomeUsuario2 == "ofelizcmadmin" || nomeUsuario2 == "helder.silva")
                    {
                        guna2ImageButton7.Visible = true;
                        ButtonExportExcelTodas.Visible = true;
                        guna2ImageButton5.Visible = true;
                        guna2ImageButton4.Visible = true;
                        guna2ImageButton3.Visible = true;
                        guna2ImageButton6.Visible = true;
                        guna2ImageButton2.Visible = true;
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

        private void FormatacaodosGraficos()
        {
              foreach (var point in chartCircle.Series["Percentagens"].Points)
                  {
                    point.Label = point.YValues[0].ToString("F1");
                  }
            
             foreach (var point in chartTotalPercentagem.Series["% Total"].Points)
                  {
                    point.Label = point.YValues[0].ToString("F1"); 
                  }

        }
                
        private void ComunicaBDparaTabelaOrcamentacao()
        {
            ComunicaBD comunicaBD = new ComunicaBD();
            try
            {
                comunicaBD.ConectarBD();
                string query = "SELECT Id, [Ano de fecho], [Numero da Obra], [Nome da Obra], [Preparador Responsavel], Tipologia, [KG Estrutura], [Horas Estrutura], [Valor Estrutura], [KG/Euro Estrutura], [Horas Revestimentos], [Valor Revestimentos], [Total Horas], [Total Valor] FROM dbo.Orçamentação";
                DataTable dataTable = comunicaBD.Procurarbd(query);
                foreach (DataRow row in dataTable.Rows)
                {
                    for (int i = 0; i < dataTable.Columns.Count; i++)
                    {
                        if (row[i] != DBNull.Value && row[i] is string)
                        {
                            row[i] = ((string)row[i]).Trim();  
                        }
                    }
                }

                DataGridViewOrcamentacaoObras.DataSource = dataTable;
                DataGridViewOrcamentacaoObras.ClearSelection();
                DataGridViewOrcamentacaoObras.ScrollBars = ScrollBars.Horizontal;
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

                //Preço da hora 
                string LabelPrecoOrcamen = LabelPrecoOrcamentado.Text;
                double precoOrcamentadoDouble = Convert.ToDouble(LabelPrecoOrcamen);

                //Horas Estrutura 
                double horasEstruturaDouble = Convert.ToDouble(HorasEstrutura);
                string horasEstruturaString = horasEstruturaDouble.ToString("F1");
                double horasEstruturaDouble2 = Convert.ToDouble(horasEstruturaString);
                double ValorEstrutura = horasEstruturaDouble2 * precoOrcamentadoDouble;

                //Valor da Estrutura
                string valorEstruturaString = ValorEstrutura.ToString("F2");
                double valorEstruturaDouble = Convert.ToDouble(valorEstruturaString);

                //Valor da KG/Estrutura
                double kgEstruturaDouble = Convert.ToDouble(KgEstrutura);
                double kgEuroEstrutura = valorEstruturaDouble / kgEstruturaDouble;
                double kgEuroEstruturaArredondado = Math.Round(kgEuroEstrutura, 2);
                string KgEuroEstruturaString = kgEuroEstruturaArredondado.ToString("F2");

                //Horas dos Revestimentos
                double HorasRevestimentosDouble = Convert.ToDouble(HorasRevestimentos);
                double ValorRevestimentos = HorasRevestimentosDouble * precoOrcamentadoDouble;
                double ValorRevestimentosArredondado = Math.Round(ValorRevestimentos, 2);
                string valorRevestimentosString = ValorRevestimentosArredondado.ToString("F2");

                //Total
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

                //Preço da hora 
                string LabelPrecoOrcamen = LabelPrecoOrcamentado.Text;
                double precoOrcamentadoDouble = Convert.ToDouble(LabelPrecoOrcamen);

                //Valor da Estrutura
                double horasEstruturaDouble = Convert.ToDouble(HorasEstrutura);
                double ValorEstrutura = horasEstruturaDouble * precoOrcamentadoDouble;
                string valorEstruturaString = ValorEstrutura.ToString("F2");

                //Valor da KG/Estrutura
                double valorEstruturaDouble = Convert.ToDouble(valorEstruturaString);
                double kgEstruturaDouble = Convert.ToDouble(KgEstrutura);
                double kgEuroEstrutura = valorEstruturaDouble / kgEstruturaDouble;
                double kgEuroEstruturaArredondado = Math.Round(kgEuroEstrutura, 2);
                string KgEuroEstruturaString = kgEuroEstruturaArredondado.ToString("F2");

                //Horas dos Revestimentos
                double HorasRevestimentosDouble = Convert.ToDouble(HorasRevestimentos);
                double ValorRevestimentos = HorasRevestimentosDouble * precoOrcamentadoDouble;
                double ValorRevestimentosArredondado = Math.Round(ValorRevestimentos, 2);
                string valorRevestimentosString = ValorRevestimentosArredondado.ToString("F2");

                //Total
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

        private void ComunicaBDparaTabelaOrcamentacaoTotal()
        {
            ComunicaBD comunicaBD = new ComunicaBD();
            try
            {
                comunicaBD.ConectarBD();

                string query = "SELECT ID, [Total KG Estrutura Orc], [Total Horas Estrutura Orc], [Total Valor Estrutura Orc], [Total KG/Euro Estrutura Orc], [Total Horas Revestimentos Orc], [Total Valor Revestimentos Orc], [Total Horas Orc], [Total Valor Orc]  FROM dbo.TotalObras";

                DataTable dataTable = comunicaBD.Procurarbd(query);

                foreach (DataRow row in dataTable.Rows)
                {
                    for (int i = 0; i < dataTable.Columns.Count; i++)
                    {
                        if (row[i] != DBNull.Value && row[i] is string)
                        {
                            row[i] = ((string)row[i]).Trim();
                        }
                    }
                }
                DataGridViewOrcamentacaoObrasTotal.DataSource = dataTable;
                DataGridViewOrcamentacaoObrasTotal.ClearSelection();
                DataGridViewOrcamentacaoObrasTotal.Columns["Id"].Visible = false;
                DataGridViewOrcamentacaoObrasTotal.ReadOnly = true;
                DataGridViewOrcamentacaoObrasTotal.ColumnHeadersVisible = false;
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

        private string SubstituirSeparadorDecimal(string valor)
        {
            return valor.Replace(",", ".");
        }
              
        private void ComunicaBDparaTabelaReal()
        {
            ComunicaBD comunicaBD = new ComunicaBD();
            try
            {
                comunicaBD.ConectarBD();

                string query = "SELECT ID, [Ano de fecho], [Numero da Obra], Tipologia, [KG Estrutura], [Horas Estrutura], [Valor Estrutura], [KG/Euro Estrutura], [Percentagem Estrutura], [Horas Revestimentos], [Valor Revestimentos] , [Percentagem Revestimentos], [Horas Aprovação], [Valor Aprovação], [Percentagem Aprovação], [Horas Alterações], [Valor Alterações], [Percentagem Alterações], [Horas Fabrico], [Valor Fabrico], [Percentagem Fabrico], [Horas Soldadura], [Valor Soldadura], [Percentagem Soldadura], [Horas Montagem], [Valor Montagem], [Percentagem Montagem], [Horas Diversos], [Valor Diversos], [Percentagem Diversos] , [Comentario Diversos], [Total Horas], [Total Valor] FROM dbo.RealObras";

                DataTable dataTable = comunicaBD.Procurarbd(query);

                foreach (DataRow row in dataTable.Rows)
                {
                    for (int i = 0; i < dataTable.Columns.Count; i++)
                    {
                        if (row[i] != DBNull.Value && row[i] is string)
                        {
                            row[i] = ((string)row[i]).Trim();
                        }
                    }
                }
                DataGridViewRealObras.DataSource = dataTable;
                DataGridViewRealObras.ClearSelection();
                DataGridViewRealObras.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
                DataGridViewRealObras.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
                DataGridViewRealObras.Columns["Id"].Visible = false;
                DataGridViewRealObras.Columns["Numero da Obra"].Visible = false;
                DataGridViewRealObras.Columns["Ano de fecho"].Visible = false;
                DataGridViewRealObras.Columns["Tipologia"].Visible = false;
                DataGridViewRealObras.ReadOnly = true;
                ApplyColumnColors();
                DataGridViewRealObras.Columns["KG Estrutura"].Width = 90;
                DataGridViewRealObras.ScrollBars = ScrollBars.Horizontal;
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

        private void ApplyColumnColors()
        {
            string[] colunasEstrutura = new string[] { "KG Estrutura", "Horas Estrutura", "Valor Estrutura", "KG/Euro Estrutura", "Percentagem Estrutura" };
            Color estruturaColor = Color.FromArgb(105, 105, 105);

            string[] colunasRevestimento = new string[] { "Horas Revestimentos", "Valor Revestimentos", "Percentagem Revestimentos" };
            Color revestimentoColor = Color.FromArgb(128, 128, 128);

            string[] colunasAprovação = new string[] { "Horas Aprovação", "Valor Aprovação", "Percentagem Aprovação" };
            Color aprovacaoColor = Color.FromArgb(169, 169, 169);

            string[] colunasAlterações = new string[] { "Horas Alterações", "Valor Alterações", "Percentagem Alterações" };
            Color alteracoesColor = Color.FromArgb(105, 105, 105);

            string[] colunasFabrico = new string[] { "Horas Fabrico", "Valor Fabrico", "Percentagem Fabrico" };
            Color fabricoColor = Color.FromArgb(128, 128, 128);

            string[] colunasSoldadura = new string[] { "Horas Soldadura", "Valor Soldadura", "Percentagem Soldadura" };
            Color soldaduraColor = Color.FromArgb(169, 169, 169);

            string[] colunasMontagem = new string[] { "Horas Montagem", "Valor Montagem", "Percentagem Montagem" };
            Color montagemColor = Color.FromArgb(105, 105, 105);

            string[] colunasDiversos = new string[] { "Horas Diversos", "Valor Diversos", "Percentagem Diversos", "Comentario Diversos" };
            Color diversosColor = Color.FromArgb(128, 128, 128);

            string[] colunasTotal = new string[] { "Total Horas", "Total Valor" };
            Color totalColor = Color.FromArgb(169, 169, 169);

            ApplyColumnStyle(colunasEstrutura, estruturaColor);
            ApplyColumnStyle(colunasRevestimento, revestimentoColor);
            ApplyColumnStyle(colunasAprovação, aprovacaoColor);
            ApplyColumnStyle(colunasAlterações, alteracoesColor);
            ApplyColumnStyle(colunasFabrico, fabricoColor);
            ApplyColumnStyle(colunasSoldadura, soldaduraColor);
            ApplyColumnStyle(colunasMontagem, montagemColor);
            ApplyColumnStyle(colunasDiversos, diversosColor);
            ApplyColumnStyle(colunasTotal, totalColor);
        }

        private void ApplyColumnStyle(string[] columns, Color headerColor)
        {
            foreach (var coluna in columns)
            {
                if (DataGridViewRealObras.Columns[coluna] != null)
                {
                    DataGridViewRealObras.Columns[coluna].HeaderCell.Style.BackColor = headerColor;
                    DataGridViewRealObras.Columns[coluna].HeaderCell.Style.ForeColor = Color.Black;

                    DataGridViewRealObras.Columns[coluna].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                }
            }
        }
              
       private void DataGridViewRealObras_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
            var columnName = DataGridViewRealObras.Columns[e.ColumnIndex].Name;

            string[] colunasEstrutura = new string[] { "KG Estrutura", "Horas Estrutura", "Valor Estrutura", "KG/Euro Estrutura", "Percentagem Estrutura" };
            string[] colunasRevestimento = new string[] { "Horas Revestimentos", "Valor Revestimentos", "Percentagem Revestimentos" };
            string[] colunasAprovação = new string[] { "Horas Aprovação", "Valor Aprovação", "Percentagem Aprovação" };
            string[] colunasAlterações = new string[] { "Horas Alterações", "Valor Alterações", "Percentagem Alterações" };
            string[] colunasFabrico = new string[] { "Horas Fabrico", "Valor Fabrico", "Percentagem Fabrico" };
            string[] colunasSoldadura = new string[] { "Horas Soldadura", "Valor Soldadura", "Percentagem Soldadura" };
            string[] colunasMontagem = new string[] { "Horas Montagem", "Valor Montagem", "Percentagem Montagem" };
            string[] colunasDiversos = new string[] { "Horas Diversos", "Valor Diversos", "Percentagem Diversos", "Comentario Diversos" };
            string[] colunasTotal = new string[] { "Total Horas", "Total Valor" };

            if (colunasEstrutura.Contains(columnName)) ApplyColumnStyle(colunasEstrutura, Color.FromArgb(244, 164, 96));
            else if (colunasRevestimento.Contains(columnName)) ApplyColumnStyle(colunasRevestimento, Color.FromArgb(250, 128, 114));
            else if (colunasAprovação.Contains(columnName)) ApplyColumnStyle(colunasAprovação, Color.FromArgb(250, 250, 210));
            else if (colunasAlterações.Contains(columnName)) ApplyColumnStyle(colunasAlterações, Color.FromArgb(135, 206, 235));
            else if (colunasFabrico.Contains(columnName)) ApplyColumnStyle(colunasFabrico, Color.FromArgb(107, 142, 35));
            else if (colunasSoldadura.Contains(columnName)) ApplyColumnStyle(colunasSoldadura, Color.FromArgb(189, 83, 107));
            else if (colunasMontagem.Contains(columnName)) ApplyColumnStyle(colunasMontagem, Color.FromArgb(169, 169, 169));
            else if (colunasDiversos.Contains(columnName)) ApplyColumnStyle(colunasDiversos, Color.FromArgb(30, 144, 255));
            else if (colunasTotal.Contains(columnName)) ApplyColumnStyle(colunasTotal, Color.FromArgb(112, 128, 144));
        }

        private void ComunicaBDparaTabelaRealTotais()
        {
            ComunicaBD comunicaBD = new ComunicaBD();
            try
            {
                comunicaBD.ConectarBD();

                string query = "SELECT ID, [Total KG Estrutura Real], [Total Horas Estrutura Real], [Total Valor Estrutura Real], [Total KG/Euro Estrutura Real], [Percentagem Estrutura Real], [Total Horas Revestimentos Real], [Total Valor Revestimentos Real], [Percentagem Revestimentos Real]," +
                    " [Total Horas Aprovacao Real] ,[Total Valor Aprovacao Real], [Percentagem Aprovacao Real], [Total Horas Alteracoes Real], [Total Valor Alteracoes Real], [Percentagem Alteracoes Real], [Total Horas Fabrico Real], [Total Valor Fabrico Real], [Percentagem Fabrico Real], [Total Horas Soldadura Real], " +
                    " [Total Valor Soldadura Real], [Percentagem Soldadura Real], [Total Horas Montagem Real], [Total Valor Montagem Real], [Percentagem Montagem Real], [Total Horas Diversos Real], [Total Valor Diversos Real], [Percentagem Diversos Real],  [Comentario Diversos Real]," +
                    " [Total Horas Real], [Total Valor Real]  FROM dbo.TotalObras";

                DataTable dataTable = comunicaBD.Procurarbd(query);

                foreach (DataRow row in dataTable.Rows)
                {
                    for (int i = 0; i < dataTable.Columns.Count; i++)
                    {
                        if (row[i] != DBNull.Value && row[i] is string)
                        {
                            row[i] = ((string)row[i]).Trim();
                        }
                    }
                }
                DataGridViewRealObrasTotal.DataSource = dataTable;
                DataGridViewRealObrasTotal.ClearSelection();
                DataGridViewRealObrasTotal.Columns["Id"].Visible = false;
                DataGridViewRealObrasTotal.ReadOnly = true;
                DataGridViewRealObrasTotal.ColumnHeadersVisible = false;
                DataGridViewRealObrasTotal.ScrollBars = ScrollBars.Horizontal;


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

                string query = "SELECT ID, [Ano de fecho], [Numero da Obra], Tipologia, [Total Horas], [Total Valor], [Percentagem Total], [Dias de Preparação] FROM dbo.ConclusaoObras";

                DataTable dataTable = comunicaBD.Procurarbd(query);

                foreach (DataRow row in dataTable.Rows)
                {
                    for (int i = 0; i < dataTable.Columns.Count; i++)
                    {
                        if (row[i] != DBNull.Value && row[i] is string)
                        {
                            row[i] = ((string)row[i]).Trim();
                        }
                    }
                }
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
            catch (Exception ex)
            {
                MessageBox.Show("Erro ao conectar à base de dados: " + ex.Message);
            }
            finally
            {
                comunicaBD.DesonectarBD();
            }
        }

        private void ComunicaBDparaTabelaConclusaoTotal()
        {
            ComunicaBD comunicaBD = new ComunicaBD();
            try
            {
                comunicaBD.ConectarBD();

                string query = "SELECT ID, [Total Horas Concl], [Total Valor Concl], [Percentagem Total Concl], [Dias de Preparacao Concl] FROM dbo.TotalObras";

                DataTable dataTable = comunicaBD.Procurarbd(query);

                foreach (DataRow row in dataTable.Rows)
                {
                    for (int i = 0; i < dataTable.Columns.Count; i++)
                    {
                        if (row[i] != DBNull.Value && row[i] is string)
                        {
                            row[i] = ((string)row[i]).Trim();
                        }
                    }
                }
                DataGridViewConclusaoObrasTotal.DataSource = dataTable;
                DataGridViewConclusaoObrasTotal.ClearSelection();
                DataGridViewConclusaoObrasTotal.Columns["Id"].Visible = false;
                DataGridViewConclusaoObrasTotal.ReadOnly = true;
                DataGridViewConclusaoObrasTotal.ColumnHeadersVisible = false;
                DataGridViewConclusaoObras.Columns["Total Horas"].Width = 90;
                DataGridViewConclusaoObras.Columns["Total Valor"].Width = 90;
                DataGridViewConclusaoObras.Columns["Percentagem Total"].Width = 70;
                DataGridViewConclusaoObras.Columns["Dias de Preparação"].Width = 70;

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

        private void guna2Button1_Click_1(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(TextBoxTarefaAdd.Text))
            {
                MessageBox.Show("Por favor, insira um valor.", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                try
                {
                    double precoOrcamentado = Convert.ToDouble(TextBoxTarefaAdd.Text);

                    LabelPrecoOrcamentado.Text = precoOrcamentado.ToString("F2");
                }
                catch (FormatException)
                {
                    MessageBox.Show("Por favor, insira um valor válido numérico. é com \",\" e nao com \". \"", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private bool isSynchronizingSelection = false;

        private void DataGridViewRealObras_SelectionChanged(object sender, EventArgs e)
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

        private void DataGridViewConclusaoObras_SelectionChanged(object sender, EventArgs e)
        {
            if (isSynchronizingSelection) return;

            if (DataGridViewConclusaoObras.SelectedRows.Count > 0)
            {
                int selectedIndex = DataGridViewConclusaoObras.SelectedRows[0].Index;

                isSynchronizingSelection = true;

                if (DataGridViewRealObras.Rows.Count > selectedIndex)
                {
                    DataGridViewRealObras.ClearSelection();
                    DataGridViewRealObras.Rows[selectedIndex].Selected = true;
                }

                if (DataGridViewOrcamentacaoObras.Rows.Count > selectedIndex)
                {
                    DataGridViewOrcamentacaoObras.ClearSelection();
                    DataGridViewOrcamentacaoObras.Rows[selectedIndex].Selected = true;
                }

                isSynchronizingSelection = false;
            }
        }

        private bool isSynchronizingScroll = false;

        private void Guna2VScrollBar_Scroll(object sender, ScrollEventArgs e)
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

            guna2HScrollBar1.Value = DataGridViewRealObras.HorizontalScrollingOffset;

            isSynchronizingScroll = false;
        }

        private void DataGridViewRealObrasTotal_ScrollHorizontal(object sender, ScrollEventArgs e)
        {
            if (isSynchronizingScroll) return;

            isSynchronizingScroll = true;

            guna2HScrollBar1.Value = DataGridViewRealObrasTotal.HorizontalScrollingOffset;

            isSynchronizingScroll = false;
        }

        private void Guna2HScrollBar_Scroll(object sender, ScrollEventArgs e)
        {
            DataGridViewRealObras.HorizontalScrollingOffset = e.NewValue;
            DataGridViewRealObrasTotal.HorizontalScrollingOffset = e.NewValue;
        }

        private void InicializarSincronizacaoDeRolagem()
        {
            DataGridViewRealObras.Scroll += DataGridViewRealObras_Scroll;
            DataGridViewOrcamentacaoObras.Scroll += DataGridViewOrcamentacaoObras_Scroll;
            DataGridViewConclusaoObras.Scroll += DataGridViewConclusaoObras_Scroll;

            DataGridViewRealObras.Scroll += DataGridViewRealObras_ScrollHorizontal;
            DataGridViewRealObrasTotal.Scroll += DataGridViewRealObrasTotal_ScrollHorizontal;

            guna2HScrollBar1.Scroll += Guna2HScrollBar_Scroll;

            guna2VScrollBar1.Scroll += Guna2VScrollBar_Scroll;

            guna2HScrollBar1.Maximum = Math.Max(DataGridViewRealObras.HorizontalScrollingOffset, DataGridViewRealObrasTotal.HorizontalScrollingOffset);
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

        private void guna2Button1_Click(object sender, EventArgs e)
        {
            guna2ContainerControl2.Visible = !guna2ContainerControl2.Visible;
            label5.Visible = !label5.Visible;
            pictureBox7.Visible = !pictureBox7.Visible;
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
                                            {  }
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
                            PercentagemTotal = ( TotalHorasReal / TotalHorasOrcamentacao) * 100;
                            PercentagemTotal = Math.Round(PercentagemTotal, 1);
                            string PercentagemT = PercentagemTotal.ToString("0.0") + "%";
                            DiasPreparar = TotalHorasResultado / 8 ;
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
                                            {        }
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
                            string KgEstrutura = guna2TextBox1.Text;
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

        private void guna2Button5_Click(object sender, EventArgs e)
        {
            AtualizarTabelaRealnaBdManual();
            AtualizarDados();
            CalcularTabelas();
        }

        private void guna2Button3_Click(object sender, EventArgs e)
        {
            ComunicarTabelas();
            AtualizarDados();
            CalcularTabelas();
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
                if (kgEuro > 0) linhaCountKGEuro++;  // Contar a linha se KG/Euro Estrutura for válido

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

        public void CarregarGraficoObras()
        {
            ComunicaBD BD = new ComunicaBD();
            BD.ConectarBD();
            int obraId = 1;
            string query = @"
            SELECT [Total Horas Orc], [Total Horas Real]
            FROM dbo.TotalObras
            WHERE ID = @ID";

            using (SqlCommand cmd = new SqlCommand(query, BD.GetConnection()))
            {
                cmd.Parameters.AddWithValue("@ID", obraId);

                using (SqlDataReader reader = cmd.ExecuteReader())
                {
                    chartTotalHoras1.Series["Orçamentação"].Points.Clear();
                    chartTotalHoras1.Series["Real"].Points.Clear();

                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {

                            double totalHorasOrc;
                            double totalHorasReal;

                            string totalHorasOrcStr = reader["Total Horas Orc"].ToString().Replace("h", "").Trim();
                            string totalHorasRealStr = reader["Total Horas Real"].ToString().Replace("h", "").Trim();

                            bool podeConverterOrc = double.TryParse(totalHorasOrcStr, out totalHorasOrc);
                            bool podeConverterReal = double.TryParse(totalHorasRealStr, out totalHorasReal);


                            if (podeConverterOrc)
                            {
                                chartTotalHoras1.Series["Orçamentação"].Points.AddY(totalHorasOrc);
                            }
                            else
                            {
                                chartTotalHoras1.Series["Orçamentação"].Points.AddY(0);
                            }

                            if (podeConverterReal)
                            {
                                chartTotalHoras1.Series["Real"].Points.AddY(totalHorasReal);
                            }
                            else
                            {
                                chartTotalHoras1.Series["Real"].Points.AddY(0);
                            }
                        }
                    }
                }
            }
            BD.DesonectarBD();
        }

        public void CarregarGraficoObras2()
        {            
            ComunicaBD BD = new ComunicaBD();
            BD.ConectarBD();

            string queryOrc = @"
                                SELECT [Total Horas], [Numero da Obra]
                                FROM dbo.Orçamentação
                                ORDER BY ID ASC"; 

            string queryReal = @"
                                SELECT [Total Horas], [Numero da Obra]
                                FROM dbo.RealObras
                                ORDER BY ID ASC";

            

            chartObrasHoras.Series["Orçamentação"].Points.Clear();
            chartObrasHoras.Series["Real"].Points.Clear();
            using (SqlCommand cmd = new SqlCommand(queryOrc, BD.GetConnection()))
            {
                using (SqlDataReader reader = cmd.ExecuteReader())
                {
                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            double totalHorasOrc;
                            string totalHorasOrcStr = reader["Total Horas"].ToString().Replace("h", "").Trim();

                            totalHorasOrcStr = totalHorasOrcStr.Replace(".", ",");

                            bool podeConverterOrc = double.TryParse(totalHorasOrcStr, out totalHorasOrc);
                            string numeroDaObraOrc = reader["Numero da Obra"].ToString();

                            if (podeConverterOrc)
                            {
                                chartObrasHoras.Series["Orçamentação"].Points.AddXY(numeroDaObraOrc, totalHorasOrc);
                            }
                            else
                            {
                                chartObrasHoras.Series["Orçamentação"].Points.AddXY(numeroDaObraOrc, 0);

                            }
                        }
                    }
                }
            }

            using (SqlCommand cmd = new SqlCommand(queryReal, BD.GetConnection()))
            {
                using (SqlDataReader reader = cmd.ExecuteReader())
                {
                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            double totalHorasReal;
                            string totalHorasRealStr = reader["Total Horas"].ToString().Replace("h", "").Trim();

                            totalHorasRealStr = totalHorasRealStr.Replace(".", ",");

                            bool podeConverterReal = double.TryParse(totalHorasRealStr, out totalHorasReal);
                            string numeroDaObraReal = reader["Numero da Obra"].ToString();

                            if (podeConverterReal)
                            {
                                chartObrasHoras.Series["Real"].Points.AddXY(numeroDaObraReal, totalHorasReal);
                            }
                            else
                            {
                                MessageBox.Show($"Erro na conversão de Total Horas Real: {totalHorasRealStr}");
                                chartObrasHoras.Series["Real"].Points.AddXY(numeroDaObraReal, 0); 
                            }
                        }
                    }
                }
            }

            BD.DesonectarBD();
        }

        public void CarregarGraficoObrasvalor()
        {
            ComunicaBD BD = new ComunicaBD();
            BD.ConectarBD();

            string queryOrc = @"
                        SELECT [Total Valor], [Numero da Obra]
                        FROM dbo.Orçamentação
                        ORDER BY ID ASC";

            string queryReal = @"
                        SELECT [Total Valor], [Numero da Obra]
                        FROM dbo.RealObras
                        ORDER BY ID ASC";

            chartTotalValorObras.Series["Orçamentação"].Points.Clear();
            chartTotalValorObras.Series["Real"].Points.Clear();
            using (SqlCommand cmd = new SqlCommand(queryOrc, BD.GetConnection()))
            {
                using (SqlDataReader reader = cmd.ExecuteReader())
                {
                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            double totalValorOrc;
                            string totalValorOrcStr = reader["Total Valor"].ToString().Replace("€", "").Trim();

                            totalValorOrcStr = totalValorOrcStr.Replace(".", ",");

                            bool podeConverterOrc = double.TryParse(totalValorOrcStr, out totalValorOrc);
                            string numeroDaObraOrc = reader["Numero da Obra"].ToString();

                            if (podeConverterOrc)
                            {
                                chartTotalValorObras.Series["Orçamentação"].Points.AddXY(numeroDaObraOrc, totalValorOrc);
                            }
                            else
                            {
                                chartTotalValorObras.Series["Orçamentação"].Points.AddXY(numeroDaObraOrc, 0);
                            }
                        }
                    }
                }
            }

            using (SqlCommand cmd = new SqlCommand(queryReal, BD.GetConnection()))
            {
                using (SqlDataReader reader = cmd.ExecuteReader())
                {
                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            double totalValorReal;
                            string totalValorRealStr = reader["Total Valor"].ToString().Replace("€", "").Trim();

                            totalValorRealStr = totalValorRealStr.Replace(".", ",");

                            bool podeConverterReal = double.TryParse(totalValorRealStr, out totalValorReal);
                            string numeroDaObraReal = reader["Numero da Obra"].ToString();

                            if (podeConverterReal)
                            {
                                chartTotalValorObras.Series["Real"].Points.AddXY(numeroDaObraReal, totalValorReal);
                            }
                            else
                            {
                                MessageBox.Show($"Erro na conversão de Total Horas Real: {totalValorRealStr}");
                                chartTotalValorObras.Series["Real"].Points.AddXY(numeroDaObraReal, 0);
                            }
                        }
                    }
                }
            }

            BD.DesonectarBD();
        }

        public void CarregarGraficoObrasPercentagem()
        {
            ComunicaBD BD = new ComunicaBD();
            BD.ConectarBD();
                        
            string queryReal = @"
                        SELECT [Percentagem Total], [Numero da Obra]
                        FROM dbo.ConclusaoObras
                        ORDER BY ID ASC";

            chartTotalPercentagem.Series["% Total"].Points.Clear();            
            using (SqlCommand cmd = new SqlCommand(queryReal, BD.GetConnection()))
            {
                using (SqlDataReader reader = cmd.ExecuteReader())
                {
                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            double totalValorPercentagem;
                            string totalValorPercentagemStr = reader["Percentagem Total"].ToString().Replace("%", "").Trim();

                            totalValorPercentagemStr = totalValorPercentagemStr.Replace(".", ",");

                            bool podeConverterReal = double.TryParse(totalValorPercentagemStr, out totalValorPercentagem);
                            string numeroDaObraReal = reader["Numero da Obra"].ToString();

                            if (podeConverterReal)
                            {
                                chartTotalPercentagem.Series["% Total"].Points.AddXY(numeroDaObraReal, totalValorPercentagem);
                            }
                            else
                            {
                                MessageBox.Show($"Erro na conversão de Total Horas Real: {totalValorPercentagemStr}");
                                chartTotalPercentagem.Series["% Total"].Points.AddXY(numeroDaObraReal, 0);
                            }
                        }
                    }
                }
            }

            BD.DesonectarBD();
        }

        public void CarregarGraficoPiePercentagem()
        {
            ComunicaBD BD = new ComunicaBD();
            BD.ConectarBD();

            string queryReal = @"
            SELECT [Percentagem Estrutura Real], [Percentagem Revestimentos Real], [Percentagem Aprovacao Real],
                   [Percentagem Alteracoes Real], [Percentagem Fabrico Real], [Percentagem Soldadura Real], 
                   [Percentagem Montagem Real], [Percentagem Diversos Real]
            FROM dbo.TotalObras";

            chartCircle.Series["Percentagens"].Points.Clear();

            chartCircle.Series["Percentagens"].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Pie;

            chartCircle.Series["Percentagens"].IsValueShownAsLabel = true; 

            chartCircle.Legends["Legend1"].Enabled = true;
            chartCircle.Legends["Legend1"].Docking = System.Windows.Forms.DataVisualization.Charting.Docking.Top;

            using (SqlCommand cmd = new SqlCommand(queryReal, BD.GetConnection()))
            {
                using (SqlDataReader reader = cmd.ExecuteReader())
                {
                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            double estruturaReal, revestimentosReal, aprovacaoReal, alteracoesReal, fabricoReal, soldaduraReal, montagemReal, diversosReal;

                            string estruturaRealStr = reader["Percentagem Estrutura Real"].ToString().Replace("%", "").Trim();
                            string revestimentosRealStr = reader["Percentagem Revestimentos Real"].ToString().Replace("%", "").Trim();
                            string aprovacaoRealStr = reader["Percentagem Aprovacao Real"].ToString().Replace("%", "").Trim();
                            string alteracoesRealStr = reader["Percentagem Alteracoes Real"].ToString().Replace("%", "").Trim();
                            string fabricoRealStr = reader["Percentagem Fabrico Real"].ToString().Replace("%", "").Trim();
                            string soldaduraRealStr = reader["Percentagem Soldadura Real"].ToString().Replace("%", "").Trim();
                            string montagemRealStr = reader["Percentagem Montagem Real"].ToString().Replace("%", "").Trim();
                            string diversosRealStr = reader["Percentagem Diversos Real"].ToString().Replace("%", "").Trim();

                            bool podeConverterEstrutura = double.TryParse(estruturaRealStr, out estruturaReal);
                            bool podeConverterRevestimentos = double.TryParse(revestimentosRealStr, out revestimentosReal);
                            bool podeConverterAprovacao = double.TryParse(aprovacaoRealStr, out aprovacaoReal);
                            bool podeConverterAlteracoes = double.TryParse(alteracoesRealStr, out alteracoesReal);
                            bool podeConverterFabrico = double.TryParse(fabricoRealStr, out fabricoReal);
                            bool podeConverterSoldadura = double.TryParse(soldaduraRealStr, out soldaduraReal);
                            bool podeConverterMontagem = double.TryParse(montagemRealStr, out montagemReal);
                            bool podeConverterDiversos = double.TryParse(diversosRealStr, out diversosReal);

                            if (podeConverterEstrutura)
                            {
                                chartCircle.Series["Percentagens"].Points.AddY(estruturaReal);
                                chartCircle.Series["Percentagens"].Points[chartCircle.Series["Percentagens"].Points.Count - 1].LegendText = "Estrutura";
                            }
                            else
                            {
                                chartCircle.Series["Percentagens"].Points.AddY(0);
                                chartCircle.Series["Percentagens"].Points[chartCircle.Series["Percentagens"].Points.Count - 1].LegendText = "Estrutura";
                            }

                            if (podeConverterRevestimentos)
                            {
                                chartCircle.Series["Percentagens"].Points.AddY(revestimentosReal);
                                chartCircle.Series["Percentagens"].Points[chartCircle.Series["Percentagens"].Points.Count - 1].LegendText = "Revestimentos";
                            }
                            else
                            {
                                chartCircle.Series["Percentagens"].Points.AddY(0);
                                chartCircle.Series["Percentagens"].Points[chartCircle.Series["Percentagens"].Points.Count - 1].LegendText = "Revestimentos";
                            }

                            if (podeConverterAprovacao)
                            {
                                chartCircle.Series["Percentagens"].Points.AddY(aprovacaoReal);
                                chartCircle.Series["Percentagens"].Points[chartCircle.Series["Percentagens"].Points.Count - 1].LegendText = "Aprovação";
                            }
                            else
                            {
                                chartCircle.Series["Percentagens"].Points.AddY(0);
                                chartCircle.Series["Percentagens"].Points[chartCircle.Series["Percentagens"].Points.Count - 1].LegendText = "Aprovação";
                            }

                            if (podeConverterAlteracoes)
                            {
                                chartCircle.Series["Percentagens"].Points.AddY(alteracoesReal);
                                chartCircle.Series["Percentagens"].Points[chartCircle.Series["Percentagens"].Points.Count - 1].LegendText = "Alterações";
                            }
                            else
                            {
                                chartCircle.Series["Percentagens"].Points.AddY(0);
                                chartCircle.Series["Percentagens"].Points[chartCircle.Series["Percentagens"].Points.Count - 1].LegendText = "Alterações";
                            }

                            if (podeConverterFabrico)
                            {
                                chartCircle.Series["Percentagens"].Points.AddY(fabricoReal);
                                chartCircle.Series["Percentagens"].Points[chartCircle.Series["Percentagens"].Points.Count - 1].LegendText = "Fabrico";
                            }
                            else
                            {
                                chartCircle.Series["Percentagens"].Points.AddY(0);
                                chartCircle.Series["Percentagens"].Points[chartCircle.Series["Percentagens"].Points.Count - 1].LegendText = "Fabrico";
                            }

                            if (podeConverterSoldadura)
                            {
                                chartCircle.Series["Percentagens"].Points.AddY(soldaduraReal);
                                chartCircle.Series["Percentagens"].Points[chartCircle.Series["Percentagens"].Points.Count - 1].LegendText = "Soldadura";
                            }
                            else
                            {
                                chartCircle.Series["Percentagens"].Points.AddY(0);
                                chartCircle.Series["Percentagens"].Points[chartCircle.Series["Percentagens"].Points.Count - 1].LegendText = "Soldadura";
                            }

                            if (podeConverterMontagem)
                            {
                                chartCircle.Series["Percentagens"].Points.AddY(montagemReal);
                                chartCircle.Series["Percentagens"].Points[chartCircle.Series["Percentagens"].Points.Count - 1].LegendText = "Montagem";
                            }
                            else
                            {
                                chartCircle.Series["Percentagens"].Points.AddY(0);
                                chartCircle.Series["Percentagens"].Points[chartCircle.Series["Percentagens"].Points.Count - 1].LegendText = "Montagem";
                            }

                            if (podeConverterDiversos)
                            {
                                chartCircle.Series["Percentagens"].Points.AddY(diversosReal);
                                chartCircle.Series["Percentagens"].Points[chartCircle.Series["Percentagens"].Points.Count - 1].LegendText = "Diversos";
                            }
                            else
                            {
                                chartCircle.Series["Percentagens"].Points.AddY(0);
                                chartCircle.Series["Percentagens"].Points[chartCircle.Series["Percentagens"].Points.Count - 1].LegendText = "Diversos";
                            }
                        }
                    }
                }
            }
            BD.DesonectarBD();
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

            if (Directory.Exists(caminhoPasta))
            {
                string[] subpastas = Directory.GetDirectories(caminhoPasta);

                ComboBoxAnoAdd2.Items.Clear();
                string pattern = @"^[A-Za-z0-9]{4}$";
                foreach (string subpasta in subpastas)
                {
                    string nomePasta = Path.GetFileName(subpasta);

                    if (Regex.IsMatch(nomePasta, pattern))
                    {
                        ComboBoxAnoAdd2.Items.Add(nomePasta);
                    }
                }
            }
            else
            {
                MessageBox.Show("O caminho especificado não existe.");
            }
        }

        private void ButtonIniciarTarefa_Click(object sender, EventArgs e)
        {
            FiltrarTipologia();
            CarregarGraficoObras();
            AdicionarSufixosNasColunas();
            CarregarGraficoObras2Tipologia();
            CarregarGraficoObrasvalorTipologia();
            CarregarGraficoObrasPercentagemTipologia();
            CarregarGraficoHorasTipologiaeAno();
            CarregarGraficoPiePercentagemTipologiaeAno();
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

        private void guna2Button3_Click_1(object sender, EventArgs e)
        {
            FiltrarPorAno();
            AdicionarSufixosNasColunas();
            CarregarGraficoHorasTipologiaeAno();
            CarregarGraficoPiePercentagemTipologiaeAno();
            CarregarGraficoObras2Ano();
            CarregarGraficoObrasvalorAno();
            CarregarGraficoObrasPercentagemAno();
        }

        private void guna2ImageButton1_Click(object sender, EventArgs e)
        {
            panel1.Visible = !panel1.Visible;
            ComunicarTabelas();
            AdicionarSufixosNasColunas();
            if (panel1.Visible == false)
            {
                CarregarGraficos();
            }            
        }

        private void guna2ImageButton2_Click(object sender, EventArgs e)
        {
            RemoverSufixosNasColunasOracamentacao();
            VerificarEAtualizarOuSalvar();
            ComunicaBDparaTabelaOrcamentacao();
            ComunicaBDparaTabelaReal();
            ComunicaBDparaTabelaConcluido();
            ComunicaBDparaTabelaOrcamentacao();
            AdicionarSufixosNasColunas();
        }

        private void guna2ImageButton4_Click(object sender, EventArgs e)
        {
            guna2Panel3.Visible = !guna2Panel3.Visible;
        }

        private void guna2ImageButton3_Click(object sender, EventArgs e)
        {
            panel2.Visible = !panel2.Visible;
        }

        private void guna2ImageButton5_Click(object sender, EventArgs e)
        {
            guna2Panel1.Visible = !guna2Panel1.Visible;
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

        private void guna2Button4_Click(object sender, EventArgs e)
        {
            InserirAnofecho();
            ComunicarTabelas();
        }

        private void guna2Button6_Click(object sender, EventArgs e)
        {
            InserirTipologia();
            ComunicarTabelas();
        }

        private void guna2ImageButton6_Click(object sender, EventArgs e)
        {
            panel3.Visible = !panel3.Visible;
        }

        private void guna2Button7_Click(object sender, EventArgs e)
        {
            LimparAnoFecho();
            ComunicarTabelas();
        }

        private void ButtonExportExcelTodas_Click(object sender, EventArgs e)
        {
            ExportExcelRegistodeTodos();
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

        private void guna2ImageButton7_Click(object sender, EventArgs e)
        {
            ExportarTabelaValoresExcel();
        }

        public DataTable DataGridViewToDataTable(DataGridView dataGridView)
        {
            DataTable dataTable = new DataTable();

            foreach (DataGridViewColumn column in dataGridView.Columns)
            {
                if (column.Visible)  
                {
                    dataTable.Columns.Add(column.HeaderText);
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
                            dataRow[cellIndex] = cell.Value;
                            cellIndex++;
                        }
                    }

                    dataTable.Rows.Add(dataRow);
                }
            }

            return dataTable;
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

        private void AdicionarSufixosNasColunas()
        {
            foreach (DataGridViewRow row in DataGridViewOrcamentacaoObras.Rows)
            {
                // **Coluna KG Estrutura**
                if (row.Cells["KG Estrutura"].Value != DBNull.Value && row.Cells["KG Estrutura"].Value != null)
                {
                    double kgValue;
                    bool isKgNumeric = Double.TryParse(row.Cells["KG Estrutura"].Value.ToString(), out kgValue);

                    if (isKgNumeric)
                    {
                        row.Cells["KG Estrutura"].Value = $"{kgValue} kg";
                    }
                }

                // **Coluna Horas Estrutura**
                if (row.Cells["Horas Estrutura"].Value != DBNull.Value && row.Cells["Horas Estrutura"].Value != null)
                {
                    double horasValue;
                    bool isHorasNumeric = Double.TryParse(row.Cells["Horas Estrutura"].Value.ToString(), out horasValue);

                    if (isHorasNumeric)
                    {
                        row.Cells["Horas Estrutura"].Value = $"{horasValue} h";
                    }
                }

                // **Coluna Valor Estrutura**
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

                // **Coluna KG/Euro Estrutura**
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

                // **Coluna Horas Revestimentos**
                if (row.Cells["Horas Revestimentos"].Value != DBNull.Value && row.Cells["Horas Revestimentos"].Value != null)
                {
                    double horasValue;
                    bool isHorasNumeric = Double.TryParse(row.Cells["Horas Revestimentos"].Value.ToString(), out horasValue);

                    if (isHorasNumeric)
                    {
                        row.Cells["Horas Revestimentos"].Value = $"{horasValue} h";
                    }
                }

                // **Coluna Valor Revestimentos**
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

                // **Coluna Total Horas**
                if (row.Cells["Total Horas"].Value != DBNull.Value && row.Cells["Total Horas"].Value != null)
                {
                    double horasValue;
                    bool isHorasNumeric = Double.TryParse(row.Cells["Total Horas"].Value.ToString(), out horasValue);

                    if (isHorasNumeric)
                    {
                        row.Cells["Total Horas"].Value = $"{horasValue} h";
                    }
                }

                // **Coluna Total Valor**
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
                // **Coluna KG Estrutura**
                if (row.Cells["KG Estrutura"].Value != DBNull.Value && row.Cells["KG Estrutura"].Value != null)
                {
                    double kgValue;
                    bool isKgNumeric = Double.TryParse(row.Cells["KG Estrutura"].Value.ToString(), out kgValue);

                    if (isKgNumeric)
                    {
                        row.Cells["KG Estrutura"].Value = $"{kgValue} kg";
                    }
                }

                // **Coluna Horas Estrutura**
                if (row.Cells["Horas Estrutura"].Value != DBNull.Value && row.Cells["Horas Estrutura"].Value != null)
                {
                    double horasValue;
                    bool isHorasNumeric = Double.TryParse(row.Cells["Horas Estrutura"].Value.ToString(), out horasValue);

                    if (isHorasNumeric)
                    {
                        row.Cells["Horas Estrutura"].Value = $"{horasValue} h";
                    }
                }

                // **Coluna Valor Estrutura**
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

                // **Coluna KG/Euro Estrutura**
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

                // **Coluna Horas Revestimentos**
                if (row.Cells["Horas Revestimentos"].Value != DBNull.Value && row.Cells["Horas Revestimentos"].Value != null)
                {
                    double horasValue;
                    bool isHorasNumeric = Double.TryParse(row.Cells["Horas Revestimentos"].Value.ToString(), out horasValue);

                    if (isHorasNumeric)
                    {
                        row.Cells["Horas Revestimentos"].Value = $"{horasValue} h";
                    }
                }

                // **Coluna Valor Revestimentos**
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

                // **Coluna Horas Aprovação**
                if (row.Cells["Horas Aprovação"].Value != DBNull.Value && row.Cells["Horas Aprovação"].Value != null)
                {
                    double horasValue;
                    bool isHorasNumeric = Double.TryParse(row.Cells["Horas Aprovação"].Value.ToString(), out horasValue);

                    if (isHorasNumeric)
                    {
                        row.Cells["Horas Aprovação"].Value = $"{horasValue} h";
                    }
                }

                // **Coluna Valor Aprovação**
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

                // **Coluna Horas Alterações**
                if (row.Cells["Horas Alterações"].Value != DBNull.Value && row.Cells["Horas Alterações"].Value != null)
                {
                    double horasValue;
                    bool isHorasNumeric = Double.TryParse(row.Cells["Horas Alterações"].Value.ToString(), out horasValue);

                    if (isHorasNumeric)
                    {
                        row.Cells["Horas Alterações"].Value = $"{horasValue} h";
                    }
                }

                // **Coluna Valor Alterações**
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

                // **Coluna Horas Fabrico**
                if (row.Cells["Horas Fabrico"].Value != DBNull.Value && row.Cells["Horas Fabrico"].Value != null)
                {
                    double horasValue;
                    bool isHorasNumeric = Double.TryParse(row.Cells["Horas Fabrico"].Value.ToString(), out horasValue);

                    if (isHorasNumeric)
                    {
                        row.Cells["Horas Fabrico"].Value = $"{horasValue} h";
                    }
                }

                // **Coluna Valor Fabrico**
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

                // **Coluna Horas Soldadura**
                if (row.Cells["Horas Soldadura"].Value != DBNull.Value && row.Cells["Horas Soldadura"].Value != null)
                {
                    double horasValue;
                    bool isHorasNumeric = Double.TryParse(row.Cells["Horas Soldadura"].Value.ToString(), out horasValue);

                    if (isHorasNumeric)
                    {
                        row.Cells["Horas Soldadura"].Value = $"{horasValue} h";
                    }
                }

                // **Coluna Valor Soldadura**
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

                // **Coluna Horas Montagem**
                if (row.Cells["Horas Montagem"].Value != DBNull.Value && row.Cells["Horas Montagem"].Value != null)
                {
                    double horasValue;
                    bool isHorasNumeric = Double.TryParse(row.Cells["Horas Montagem"].Value.ToString(), out horasValue);

                    if (isHorasNumeric)
                    {
                        row.Cells["Horas Montagem"].Value = $"{horasValue} h";
                    }
                }

                // **Coluna Valor Montagem**
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

                // **Coluna Horas Diversos**
                if (row.Cells["Horas Diversos"].Value != DBNull.Value && row.Cells["Horas Diversos"].Value != null)
                {
                    double horasValue;
                    bool isHorasNumeric = Double.TryParse(row.Cells["Horas Diversos"].Value.ToString(), out horasValue);

                    if (isHorasNumeric)
                    {
                        row.Cells["Horas Diversos"].Value = $"{horasValue} h";
                    }
                }

                // **Coluna Valor Diversos**
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

                // **Coluna Total Horas**
                if (row.Cells["Total Horas"].Value != DBNull.Value && row.Cells["Total Horas"].Value != null)
                {
                    double horasValue;
                    bool isHorasNumeric = Double.TryParse(row.Cells["Total Horas"].Value.ToString(), out horasValue);

                    if (isHorasNumeric)
                    {
                        row.Cells["Total Horas"].Value = $"{horasValue} h";
                    }
                }

                // **Coluna Total Valor**
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
                // **Coluna Total Horas**
                if (row.Cells["Total Horas"].Value != DBNull.Value && row.Cells["Total Horas"].Value != null)
                {
                    double horasValue;
                    bool isHorasNumeric = Double.TryParse(row.Cells["Total Horas"].Value.ToString(), out horasValue);

                    if (isHorasNumeric)
                    {
                        row.Cells["Total Horas"].Value = $"{horasValue} h";
                    }
                }

                // **Coluna Total Valor**
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

                // **Coluna Dias de Preparação**
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

        private void RemoverSufixosNasColunasOracamentacao()
        {
            foreach (DataGridViewRow row in DataGridViewOrcamentacaoObras.Rows)
            {
                // **Coluna KG Estrutura**
                if (row.Cells["KG Estrutura"].Value != DBNull.Value && row.Cells["KG Estrutura"].Value != null)
                {
                    string kgValueStr = row.Cells["KG Estrutura"].Value.ToString();
                    kgValueStr = kgValueStr.Replace(" kg", "").Trim(); // Remove " kg" e espaços extras

                    double kgValue;
                    bool isKgNumeric = Double.TryParse(kgValueStr, out kgValue);

                    if (isKgNumeric)
                    {
                        row.Cells["KG Estrutura"].Value = kgValue; // Restaura o valor numérico
                    }
                }

                // **Coluna Horas Estrutura**
                if (row.Cells["Horas Estrutura"].Value != DBNull.Value && row.Cells["Horas Estrutura"].Value != null)
                {
                    string horasValueStr = row.Cells["Horas Estrutura"].Value.ToString();
                    horasValueStr = horasValueStr.Replace(" h", "").Trim(); // Remove " h" e espaços extras

                    double horasValue;
                    bool isHorasNumeric = Double.TryParse(horasValueStr, out horasValue);

                    if (isHorasNumeric)
                    {
                        row.Cells["Horas Estrutura"].Value = horasValue; // Restaura o valor numérico
                    }
                }

                // **Coluna Valor Estrutura**
                if (row.Cells["Valor Estrutura"].Value != DBNull.Value && row.Cells["Valor Estrutura"].Value != null)
                {
                    string valorStr = row.Cells["Valor Estrutura"].Value.ToString();
                    valorStr = valorStr.Replace(" €", "").Replace('.', ',').Trim(); // Remove " €", e troca a vírgula por ponto

                    double valorValue;
                    bool isValorNumeric = Double.TryParse(valorStr, out valorValue);

                    if (isValorNumeric)
                    {
                        row.Cells["Valor Estrutura"].Value = valorValue; // Restaura o valor numérico
                    }
                }

                // **Coluna KG/Euro Estrutura**
                if (row.Cells["KG/Euro Estrutura"].Value != DBNull.Value && row.Cells["KG/Euro Estrutura"].Value != null)
                {
                    string valorStr = row.Cells["KG/Euro Estrutura"].Value.ToString();
                    valorStr = valorStr.Replace(" €", "").Replace('.', ',').Trim(); // Remove " €", e troca a vírgula por ponto

                    double valorValue;
                    bool isValorNumeric = Double.TryParse(valorStr, out valorValue);

                    if (isValorNumeric)
                    {
                        row.Cells["KG/Euro Estrutura"].Value = valorValue; // Restaura o valor numérico
                    }
                }

                // **Coluna Horas Revestimentos**
                if (row.Cells["Horas Revestimentos"].Value != DBNull.Value && row.Cells["Horas Revestimentos"].Value != null)
                {
                    string horasValueStr = row.Cells["Horas Revestimentos"].Value.ToString();
                    horasValueStr = horasValueStr.Replace(" h", "").Trim(); // Remove " h" e espaços extras

                    double horasValue;
                    bool isHorasNumeric = Double.TryParse(horasValueStr, out horasValue);

                    if (isHorasNumeric)
                    {
                        row.Cells["Horas Revestimentos"].Value = horasValue; // Restaura o valor numérico
                    }
                }

                // **Coluna Valor Revestimentos**
                if (row.Cells["Valor Revestimentos"].Value != DBNull.Value && row.Cells["Valor Revestimentos"].Value != null)
                {
                    string valorStr = row.Cells["Valor Revestimentos"].Value.ToString();
                    valorStr = valorStr.Replace(" €", "").Replace('.', ',').Trim(); // Remove " €", e troca a vírgula por ponto

                    double valorValue;
                    bool isValorNumeric = Double.TryParse(valorStr, out valorValue);

                    if (isValorNumeric)
                    {
                        row.Cells["Valor Revestimentos"].Value = valorValue; // Restaura o valor numérico
                    }
                }

                // **Coluna Total Horas**
                if (row.Cells["Total Horas"].Value != DBNull.Value && row.Cells["Total Horas"].Value != null)
                {
                    string horasValueStr = row.Cells["Total Horas"].Value.ToString();
                    horasValueStr = horasValueStr.Replace(" h", "").Trim(); // Remove " h" e espaços extras

                    double horasValue;
                    bool isHorasNumeric = Double.TryParse(horasValueStr, out horasValue);

                    if (isHorasNumeric)
                    {
                        row.Cells["Total Horas"].Value = horasValue; // Restaura o valor numérico
                    }
                }

                // **Coluna Total Valor**
                if (row.Cells["Total Valor"].Value != DBNull.Value && row.Cells["Total Valor"].Value != null)
                {
                    string valorStr = row.Cells["Total Valor"].Value.ToString();
                    valorStr = valorStr.Replace(" €", "").Replace('.', ',').Trim(); // Remove " €", e troca a vírgula por ponto

                    double valorValue;
                    bool isValorNumeric = Double.TryParse(valorStr, out valorValue);

                    if (isValorNumeric)
                    {
                        row.Cells["Total Valor"].Value = valorValue; // Restaura o valor numérico
                    }
                }
            }
        }

        public void CarregarGraficoHorasTipologiaeAno()
        {
            double totalHorasOrcamento = 0;
            double totalHorasReal = 0;

            foreach (DataGridViewRow row in DataGridViewOrcamentacaoObras.Rows)
            {
                if (row.Cells["Total Horas"].Value != null)
                {
                    string totalHorasOrcStr = row.Cells["Total Horas"].Value.ToString().Replace("h", "").Trim();
                    totalHorasOrcStr = totalHorasOrcStr.Replace(".", ",");

                    double totalHorasOrc;
                    if (double.TryParse(totalHorasOrcStr, out totalHorasOrc))
                    {
                        totalHorasOrcamento += totalHorasOrc;
                    }
                    else
                    {
                        MessageBox.Show($"Erro ao converter o valor de 'Total Horas' na Orçamentação para o número da obra: {row.Cells["Numero da Obra"].Value}");
                    }
                }
            }

            foreach (DataGridViewRow row in DataGridViewRealObras.Rows)
            {
                if (row.Cells["Total Horas"].Value != null)
                {
                    string totalHorasRealStr = row.Cells["Total Horas"].Value.ToString().Replace("h", "").Trim();
                    totalHorasRealStr = totalHorasRealStr.Replace(".", ",");

                    double totalHorasRealAux;
                    if (double.TryParse(totalHorasRealStr, out totalHorasRealAux))
                    {
                        totalHorasReal += totalHorasRealAux;
                    }
                    else
                    {
                        MessageBox.Show($"Erro ao converter o valor de 'Total Horas' na RealObras para o número da obra: {row.Cells["Numero da Obra"].Value}");
                    }
                }
            }

            chartTotalHoras1.Series["Orçamentação"].Points.Clear();
            chartTotalHoras1.Series["Real"].Points.Clear();

            chartTotalHoras1.Series["Orçamentação"].Points.AddY(totalHorasOrcamento);
            chartTotalHoras1.Series["Real"].Points.AddY(totalHorasReal);

        }

        public void CarregarGraficoPiePercentagemTipologiaeAno()
        {
            chartCircle.Series["Percentagens"].Points.Clear();
            chartCircle.Series["Percentagens"].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Pie;
            chartCircle.Series["Percentagens"].IsValueShownAsLabel = true;
            chartCircle.Legends["Legend1"].Enabled = true;
            chartCircle.Legends["Legend1"].Docking = System.Windows.Forms.DataVisualization.Charting.Docking.Top;

            double somaEstrutura = 0, somaRevestimentos = 0, somaAprovacao = 0;
            double somaAlteracoes = 0, somaFabrico = 0, somaSoldadura = 0;
            double somaMontagem = 0, somaDiversos = 0;
            int count = 0;

            foreach (DataGridViewRow row in DataGridViewRealObras.Rows)
            {
                if (row.IsNewRow) continue;

                double estruturaReal = 0, revestimentosReal = 0, aprovacaoReal = 0, alteracoesReal = 0;
                double fabricoReal = 0, soldaduraReal = 0, montagemReal = 0, diversosReal = 0;

                bool podeConverterEstrutura = double.TryParse(row.Cells["Percentagem Estrutura"].Value.ToString().Replace("%", "").Trim(), out estruturaReal);
                bool podeConverterRevestimentos = double.TryParse(row.Cells["Percentagem Revestimentos"].Value.ToString().Replace("%", "").Trim(), out revestimentosReal);
                bool podeConverterAprovacao = double.TryParse(row.Cells["Percentagem Aprovação"].Value.ToString().Replace("%", "").Trim(), out aprovacaoReal);
                bool podeConverterAlteracoes = double.TryParse(row.Cells["Percentagem Alterações"].Value.ToString().Replace("%", "").Trim(), out alteracoesReal);
                bool podeConverterFabrico = double.TryParse(row.Cells["Percentagem Fabrico"].Value.ToString().Replace("%", "").Trim(), out fabricoReal);
                bool podeConverterSoldadura = double.TryParse(row.Cells["Percentagem Soldadura"].Value.ToString().Replace("%", "").Trim(), out soldaduraReal);
                bool podeConverterMontagem = double.TryParse(row.Cells["Percentagem Montagem"].Value.ToString().Replace("%", "").Trim(), out montagemReal);
                bool podeConverterDiversos = double.TryParse(row.Cells["Percentagem Diversos"].Value.ToString().Replace("%", "").Trim(), out diversosReal);

                if (podeConverterEstrutura) somaEstrutura += estruturaReal;
                if (podeConverterRevestimentos) somaRevestimentos += revestimentosReal;
                if (podeConverterAprovacao) somaAprovacao += aprovacaoReal;
                if (podeConverterAlteracoes) somaAlteracoes += alteracoesReal;
                if (podeConverterFabrico) somaFabrico += fabricoReal;
                if (podeConverterSoldadura) somaSoldadura += soldaduraReal;
                if (podeConverterMontagem) somaMontagem += montagemReal;
                if (podeConverterDiversos) somaDiversos += diversosReal;

                count++;
            }

            if (count > 0)
            {
                chartCircle.Series["Percentagens"].Points.AddY(somaEstrutura / count);
                chartCircle.Series["Percentagens"].Points[chartCircle.Series["Percentagens"].Points.Count - 1].LegendText = "Estrutura";

                chartCircle.Series["Percentagens"].Points.AddY(somaRevestimentos / count);
                chartCircle.Series["Percentagens"].Points[chartCircle.Series["Percentagens"].Points.Count - 1].LegendText = "Revestimentos";

                chartCircle.Series["Percentagens"].Points.AddY(somaAprovacao / count);
                chartCircle.Series["Percentagens"].Points[chartCircle.Series["Percentagens"].Points.Count - 1].LegendText = "Aprovação";

                chartCircle.Series["Percentagens"].Points.AddY(somaAlteracoes / count);
                chartCircle.Series["Percentagens"].Points[chartCircle.Series["Percentagens"].Points.Count - 1].LegendText = "Alterações";

                chartCircle.Series["Percentagens"].Points.AddY(somaFabrico / count);
                chartCircle.Series["Percentagens"].Points[chartCircle.Series["Percentagens"].Points.Count - 1].LegendText = "Fabrico";

                chartCircle.Series["Percentagens"].Points.AddY(somaSoldadura / count);
                chartCircle.Series["Percentagens"].Points[chartCircle.Series["Percentagens"].Points.Count - 1].LegendText = "Soldadura";

                chartCircle.Series["Percentagens"].Points.AddY(somaMontagem / count);
                chartCircle.Series["Percentagens"].Points[chartCircle.Series["Percentagens"].Points.Count - 1].LegendText = "Montagem";

                chartCircle.Series["Percentagens"].Points.AddY(somaDiversos / count);
                chartCircle.Series["Percentagens"].Points[chartCircle.Series["Percentagens"].Points.Count - 1].LegendText = "Diversos";
            }
            else
            {
                MessageBox.Show("Não há dados válidos no DataGridView.");
            }
        }

        public void CarregarGraficoObras2Tipologia()
        {
            string Tipologia = ComboBoxTipologiaFiltro.SelectedItem?.ToString();
           
             if (string.IsNullOrEmpty(Tipologia))
            {
                MessageBox.Show("Por favor, selecione uma tipologia.");
                return;
            }

            ComunicaBD BD = new ComunicaBD();
            BD.ConectarBD();

            string queryOrc = @"
                                SELECT [Total Horas], [Numero da Obra]
                                FROM dbo.Orçamentação
                                WHERE Tipologia = @Tipologia";

            string queryReal = @"
                                SELECT [Total Horas], [Numero da Obra]
                                FROM dbo.RealObras
                                WHERE Tipologia = @Tipologia";



            chartObrasHoras.Series["Orçamentação"].Points.Clear();
            chartObrasHoras.Series["Real"].Points.Clear();

            using (SqlCommand cmd = new SqlCommand(queryOrc, BD.GetConnection()))
            {
                cmd.Parameters.AddWithValue("@Tipologia", Tipologia);

                using (SqlDataReader reader = cmd.ExecuteReader())
                {
                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            double totalHorasOrc;
                            string totalHorasOrcStr = reader["Total Horas"].ToString().Replace("h", "").Trim();

                            totalHorasOrcStr = totalHorasOrcStr.Replace(".", ",");

                            bool podeConverterOrc = double.TryParse(totalHorasOrcStr, out totalHorasOrc);
                            string numeroDaObraOrc = reader["Numero da Obra"].ToString();

                            if (podeConverterOrc)
                            {
                                chartObrasHoras.Series["Orçamentação"].Points.AddXY(numeroDaObraOrc, totalHorasOrc);
                            }
                            else
                            {
                                chartObrasHoras.Series["Orçamentação"].Points.AddXY(numeroDaObraOrc, 0);

                            }
                        }
                    }
                }
            }

            using (SqlCommand cmd = new SqlCommand(queryReal, BD.GetConnection()))
            {
                cmd.Parameters.AddWithValue("@Tipologia", Tipologia);

                using (SqlDataReader reader = cmd.ExecuteReader())
                {
                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            double totalHorasReal;
                            string totalHorasRealStr = reader["Total Horas"].ToString().Replace("h", "").Trim();

                            totalHorasRealStr = totalHorasRealStr.Replace(".", ",");

                            bool podeConverterReal = double.TryParse(totalHorasRealStr, out totalHorasReal);
                            string numeroDaObraReal = reader["Numero da Obra"].ToString();

                            if (podeConverterReal)
                            {
                                chartObrasHoras.Series["Real"].Points.AddXY(numeroDaObraReal, totalHorasReal);
                            }
                            else
                            {
                                MessageBox.Show($"Erro na conversão de Total Horas Real: {totalHorasRealStr}");
                                chartObrasHoras.Series["Real"].Points.AddXY(numeroDaObraReal, 0);
                            }
                        }
                    }
                }
            }

            BD.DesonectarBD();
        }

        public void CarregarGraficoObrasvalorTipologia()
        {
            string Tipologia = ComboBoxTipologiaFiltro.SelectedItem?.ToString();

            ComunicaBD BD = new ComunicaBD();
            BD.ConectarBD();

            string queryOrc = @"
                        SELECT [Total Valor], [Numero da Obra]
                        FROM dbo.Orçamentação
                        WHERE Tipologia = @Tipologia";

            string queryReal = @"
                        SELECT [Total Valor], [Numero da Obra]
                        FROM dbo.RealObras
                        WHERE Tipologia = @Tipologia";


            chartTotalValorObras.Series["Orçamentação"].Points.Clear();
            chartTotalValorObras.Series["Real"].Points.Clear();

            using (SqlCommand cmd = new SqlCommand(queryOrc, BD.GetConnection()))
            {
                cmd.Parameters.AddWithValue("@Tipologia", Tipologia);

                using (SqlDataReader reader = cmd.ExecuteReader())
                {
                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            double totalValorOrc;
                            string totalValorOrcStr = reader["Total Valor"].ToString().Replace("€", "").Trim();

                            totalValorOrcStr = totalValorOrcStr.Replace(".", ",");

                            bool podeConverterOrc = double.TryParse(totalValorOrcStr, out totalValorOrc);
                            string numeroDaObraOrc = reader["Numero da Obra"].ToString();

                            if (podeConverterOrc)
                            {
                                chartTotalValorObras.Series["Orçamentação"].Points.AddXY(numeroDaObraOrc, totalValorOrc);
                            }
                            else
                            {
                                chartTotalValorObras.Series["Orçamentação"].Points.AddXY(numeroDaObraOrc, 0);
                            }
                        }
                    }
                }
            }

            using (SqlCommand cmd = new SqlCommand(queryReal, BD.GetConnection()))
            {
                cmd.Parameters.AddWithValue("@Tipologia", Tipologia);

                using (SqlDataReader reader = cmd.ExecuteReader())
                {
                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            double totalValorReal;
                            string totalValorRealStr = reader["Total Valor"].ToString().Replace("€", "").Trim();

                            totalValorRealStr = totalValorRealStr.Replace(".", ",");

                            bool podeConverterReal = double.TryParse(totalValorRealStr, out totalValorReal);
                            string numeroDaObraReal = reader["Numero da Obra"].ToString();

                            if (podeConverterReal)
                            {
                                chartTotalValorObras.Series["Real"].Points.AddXY(numeroDaObraReal, totalValorReal);
                            }
                            else
                            {
                                MessageBox.Show($"Erro na conversão de Total Horas Real: {totalValorRealStr}");
                                chartTotalValorObras.Series["Real"].Points.AddXY(numeroDaObraReal, 0);
                            }
                        }
                    }
                }
            }

            BD.DesonectarBD();
        }

        public void CarregarGraficoObrasPercentagemTipologia()
        {
            string Tipologia = ComboBoxTipologiaFiltro.SelectedItem?.ToString();

            ComunicaBD BD = new ComunicaBD();
            BD.ConectarBD();

            string queryReal = @"
                        SELECT [Percentagem Total], [Numero da Obra]
                        FROM dbo.ConclusaoObras
                        WHERE Tipologia = @Tipologia";


            chartTotalPercentagem.Series["% Total"].Points.Clear();

            using (SqlCommand cmd = new SqlCommand(queryReal, BD.GetConnection()))
            {
                cmd.Parameters.AddWithValue("@Tipologia", Tipologia);

                using (SqlDataReader reader = cmd.ExecuteReader())
                {
                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            double totalValorPercentagem;
                            string totalValorPercentagemStr = reader["Percentagem Total"].ToString().Replace("%", "").Trim();

                            totalValorPercentagemStr = totalValorPercentagemStr.Replace(".", ",");

                            bool podeConverterReal = double.TryParse(totalValorPercentagemStr, out totalValorPercentagem);
                            string numeroDaObraReal = reader["Numero da Obra"].ToString();

                            if (podeConverterReal)
                            {
                                chartTotalPercentagem.Series["% Total"].Points.AddXY(numeroDaObraReal, totalValorPercentagem);
                            }
                            else
                            {
                                MessageBox.Show($"Erro na conversão de Total Horas Real: {totalValorPercentagemStr}");
                                chartTotalPercentagem.Series["% Total"].Points.AddXY(numeroDaObraReal, 0);
                            }
                        }
                    }
                }
            }

            BD.DesonectarBD();
        }
              
        public void CarregarGraficoObras2Ano()
        {
            string Anofecho = ComboBoxAnoAdd.SelectedItem?.ToString();

            if (string.IsNullOrEmpty(Anofecho))
            {
                MessageBox.Show("Por favor, selecione uma Ano.");
                return;
            }


            ComunicaBD BD = new ComunicaBD();
            BD.ConectarBD();

            string queryOrc = @"
                                SELECT [Total Horas], [Numero da Obra]
                                FROM dbo.Orçamentação
                                WHERE [Ano de fecho] = @anofecho";

            string queryReal = @"
                                SELECT [Total Horas], [Numero da Obra]
                                FROM dbo.RealObras
                                WHERE [Ano de fecho] = @anofecho";



            chartObrasHoras.Series["Orçamentação"].Points.Clear();
            chartObrasHoras.Series["Real"].Points.Clear();

            using (SqlCommand cmd = new SqlCommand(queryOrc, BD.GetConnection()))
            {
                cmd.Parameters.AddWithValue("@anofecho", Anofecho);

                using (SqlDataReader reader = cmd.ExecuteReader())
                {
                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            double totalHorasOrc;
                            string totalHorasOrcStr = reader["Total Horas"].ToString().Replace("h", "").Trim();

                            totalHorasOrcStr = totalHorasOrcStr.Replace(".", ",");

                            bool podeConverterOrc = double.TryParse(totalHorasOrcStr, out totalHorasOrc);
                            string numeroDaObraOrc = reader["Numero da Obra"].ToString();

                            if (podeConverterOrc)
                            {
                                chartObrasHoras.Series["Orçamentação"].Points.AddXY(numeroDaObraOrc, totalHorasOrc);
                            }
                            else
                            {
                                chartObrasHoras.Series["Orçamentação"].Points.AddXY(numeroDaObraOrc, 0);

                            }
                        }
                    }
                }
            }

            using (SqlCommand cmd = new SqlCommand(queryReal, BD.GetConnection()))
            {
                cmd.Parameters.AddWithValue("@anofecho", Anofecho);

                using (SqlDataReader reader = cmd.ExecuteReader())
                {
                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            double totalHorasReal;
                            string totalHorasRealStr = reader["Total Horas"].ToString().Replace("h", "").Trim();

                            totalHorasRealStr = totalHorasRealStr.Replace(".", ",");

                            bool podeConverterReal = double.TryParse(totalHorasRealStr, out totalHorasReal);
                            string numeroDaObraReal = reader["Numero da Obra"].ToString();

                            if (podeConverterReal)
                            {
                                chartObrasHoras.Series["Real"].Points.AddXY(numeroDaObraReal, totalHorasReal);
                            }
                            else
                            {
                                MessageBox.Show($"Erro na conversão de Total Horas Real: {totalHorasRealStr}");
                                chartObrasHoras.Series["Real"].Points.AddXY(numeroDaObraReal, 0);
                            }
                        }
                    }
                }
            }

            BD.DesonectarBD();
        }

        public void CarregarGraficoObrasvalorAno()
        {
            string Anofecho = ComboBoxAnoAdd.SelectedItem?.ToString();

            ComunicaBD BD = new ComunicaBD();
            BD.ConectarBD();

            string queryOrc = @"
                        SELECT [Total Valor], [Numero da Obra]
                        FROM dbo.Orçamentação
                        WHERE [Ano de fecho] = @anofecho";

            string queryReal = @"
                        SELECT [Total Valor], [Numero da Obra]
                        FROM dbo.RealObras
                        WHERE [Ano de fecho] = @anofecho";


            chartTotalValorObras.Series["Orçamentação"].Points.Clear();
            chartTotalValorObras.Series["Real"].Points.Clear();

            using (SqlCommand cmd = new SqlCommand(queryOrc, BD.GetConnection()))
            {
                cmd.Parameters.AddWithValue("@anofecho", Anofecho);

                using (SqlDataReader reader = cmd.ExecuteReader())
                {
                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            double totalValorOrc;
                            string totalValorOrcStr = reader["Total Valor"].ToString().Replace("€", "").Trim();

                            totalValorOrcStr = totalValorOrcStr.Replace(".", ",");

                            bool podeConverterOrc = double.TryParse(totalValorOrcStr, out totalValorOrc);
                            string numeroDaObraOrc = reader["Numero da Obra"].ToString();

                            if (podeConverterOrc)
                            {
                                chartTotalValorObras.Series["Orçamentação"].Points.AddXY(numeroDaObraOrc, totalValorOrc);
                            }
                            else
                            {
                                chartTotalValorObras.Series["Orçamentação"].Points.AddXY(numeroDaObraOrc, 0);
                            }
                        }
                    }
                }
            }

            using (SqlCommand cmd = new SqlCommand(queryReal, BD.GetConnection()))
            {
                cmd.Parameters.AddWithValue("@anofecho", Anofecho);

                using (SqlDataReader reader = cmd.ExecuteReader())
                {
                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            double totalValorReal;
                            string totalValorRealStr = reader["Total Valor"].ToString().Replace("€", "").Trim();

                            totalValorRealStr = totalValorRealStr.Replace(".", ",");

                            bool podeConverterReal = double.TryParse(totalValorRealStr, out totalValorReal);
                            string numeroDaObraReal = reader["Numero da Obra"].ToString();

                            if (podeConverterReal)
                            {
                                chartTotalValorObras.Series["Real"].Points.AddXY(numeroDaObraReal, totalValorReal);
                            }
                            else
                            {
                                MessageBox.Show($"Erro na conversão de Total Horas Real: {totalValorRealStr}");
                                chartTotalValorObras.Series["Real"].Points.AddXY(numeroDaObraReal, 0);
                            }
                        }
                    }
                }
            }

            BD.DesonectarBD();
        }

        public void CarregarGraficoObrasPercentagemAno()
        {
            string Anofecho = ComboBoxAnoAdd.SelectedItem?.ToString();

            ComunicaBD BD = new ComunicaBD();
            BD.ConectarBD();

            string queryReal = @"
                        SELECT [Percentagem Total], [Numero da Obra]
                        FROM dbo.ConclusaoObras
                        WHERE [Ano de fecho] = @anofecho";


            chartTotalPercentagem.Series["% Total"].Points.Clear();

            using (SqlCommand cmd = new SqlCommand(queryReal, BD.GetConnection()))
            {
                cmd.Parameters.AddWithValue("@anofecho", Anofecho);

                using (SqlDataReader reader = cmd.ExecuteReader())
                {
                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            double totalValorPercentagem;
                            string totalValorPercentagemStr = reader["Percentagem Total"].ToString().Replace("%", "").Trim();

                            totalValorPercentagemStr = totalValorPercentagemStr.Replace(".", ",");

                            bool podeConverterReal = double.TryParse(totalValorPercentagemStr, out totalValorPercentagem);
                            string numeroDaObraReal = reader["Numero da Obra"].ToString();

                            if (podeConverterReal)
                            {
                                chartTotalPercentagem.Series["% Total"].Points.AddXY(numeroDaObraReal, totalValorPercentagem);
                            }
                            else
                            {
                                MessageBox.Show($"Erro na conversão de Total Horas Real: {totalValorPercentagemStr}");
                                chartTotalPercentagem.Series["% Total"].Points.AddXY(numeroDaObraReal, 0);
                            }
                        }
                    }
                }
            }

            BD.DesonectarBD();
        }

        private void chartObrasHoras_Click(object sender, EventArgs e)
        {
            chartObrasHoras.ChartAreas[0].AxisX.ScaleView.Size = 10;

            chartObrasHoras.ChartAreas[0].AxisX.ScrollBar.Enabled = true;
            chartObrasHoras.ChartAreas[0].AxisX.ScrollBar.IsPositionedInside = true;

            chartObrasHoras.ChartAreas[0].AxisX.ScaleView.Zoomable = true;

            chartObrasHoras.ChartAreas[0].CursorX.IsUserEnabled = true;
            chartObrasHoras.ChartAreas[0].CursorX.IsUserSelectionEnabled = true;

        }

    }
}




