using Guna.Charts.WinForms;
using Microsoft.Office.Interop.Excel;
using ServiceStack.Script;
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
using System.Windows.Forms.DataVisualization.Charting;
using DataTable = System.Data.DataTable;
using Action = System.Action;
using Font = System.Drawing.Font;
using System.Globalization;



namespace OfelizCM
{
    public partial class Frm_Dashbord : Form
    {

        public string NomeObra
        {
            set
            {
                labelNumeroObra.Text = value;
            }
        }

        public Frm_Dashbord()
        {
            InitializeComponent();
        }

        private void Frm_Dashbord_Load(object sender, EventArgs e)
        {
            CarregarPreparadorResponsavel();
            CarregarPercentagemTotal();
            AtualizarIndicadorPercentagem();
            CarregarPercentagensCircle();
            CarregarGraficoPiePercentagem();
            CarregarPercentagensOrcamento();
            CarregarPercentagensReal();
            CarregarValoresQuadro();
            AtualizarTabelaHorasPreparador();
            ConclusaoValores();
            PercentagemHorasValor();
            VerificarUsuario();
            CarregarHorasPorPreparador2();
            AtualizarImagemHoras();
            AtualizarImagemHoras2();
            AtualizarImagemHoras3();
            LabelFormatado();
            DateTimePickerInicio.Value = DateTime.Now;
            DateTimePickerConclusao.Value = DateTime.Now.AddDays(5);
            GraficoTotalhoras();
            GraficoTotalPercentagem();
            ImagemObraTomb();
            
        }

        private void CarregarPreparadorResponsavel()
        {
            ComunicaBD BD = new ComunicaBD();
            BD.ConectarBD();
            string NumeroObra = labelNumeroObra.Text;
            labelObraExcel.Text = NumeroObra;

            string preparadorResponsavel = string.Empty;
            string NomeObra = string.Empty;
            string query = @"
                    SELECT [Preparador Responsavel], [Nome da Obra]
                    FROM dbo.RealObras
                    WHERE [Numero da Obra] = @NumeroDaObra";

            using (SqlCommand cmd = new SqlCommand(query, BD.GetConnection()))
            {
                cmd.Parameters.AddWithValue("@NumeroDaObra", NumeroObra);

                using (SqlDataReader reader = cmd.ExecuteReader())
                {
                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            preparadorResponsavel = reader["Preparador Responsavel"].ToString();
                            NomeObra = reader["Nome da Obra"].ToString();
                        }

                        labelNomePreparador.Text = preparadorResponsavel;
                        labelNomeObra.Text = NomeObra;
                    }
                    else
                    {
                        labelNomePreparador.Text = "Não encontrado Preparador Responsavel";
                    }
                }
            }
            BD.DesonectarBD();
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
                            label8.Visible = true;
                            labelObraExcel.Visible = true;
                            ButtonExcel.Visible = true;
                            label15.Visible = true;
                            label13.Visible = true;
                            DateTimePickerInicio.Visible = true;
                            DateTimePickerConclusao.Visible = true;
                            ButtonExcel2.Visible = true;


                        }
                        else
                        {
                            label8.Visible = false;
                            labelObraExcel.Visible = false;
                            ButtonExcel.Visible = false;
                            label15.Visible = false;
                            label13.Visible = false;
                            DateTimePickerInicio.Visible = false;
                            DateTimePickerConclusao.Visible = false;
                            ButtonExcel2.Visible = false;

                        }
                    }
                    else
                    {

                        label8.Visible = false;
                        labelObraExcel.Visible = false;
                        ButtonExcel.Visible = false;
                        label15.Visible = false;
                        label13.Visible = false;
                        DateTimePickerInicio.Visible = false;
                        DateTimePickerConclusao.Visible = false;
                        ButtonExcel2.Visible = false;
                    }
                    string nomeUsuario2 = Properties.Settings.Default.NomeUsuario;

                    if (nomeUsuario2 == "ofelizcmadmin" || nomeUsuario2 == "helder.silva")
                    {
                        label8.Visible = true;
                        labelObraExcel.Visible = true;
                        ButtonExcel.Visible = true;
                        label15.Visible = true;
                        label13.Visible = true;
                        DateTimePickerInicio.Visible = true;
                        DateTimePickerConclusao.Visible = true;
                        ButtonExcel2.Visible = true;
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

        private void CarregarPercentagemTotal()
        {
            ComunicaBD BD = new ComunicaBD();
            BD.ConectarBD();
            string NumeroObra = labelNumeroObra.Text.Trim();

            string percentagemTotal = string.Empty;

            string query = @"
                    SELECT [Percentagem Total]
                    FROM dbo.ConclusaoObras
                    WHERE [Numero da Obra] = @NumeroDaObra";

            using (SqlCommand cmd = new SqlCommand(query, BD.GetConnection()))
            {
                cmd.Parameters.AddWithValue("@NumeroDaObra", NumeroObra);

                using (SqlDataReader reader = cmd.ExecuteReader())
                {
                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            percentagemTotal = reader["Percentagem Total"].ToString();
                        }

                        labelIndicadorPercentagem.Text = percentagemTotal + "%";
                    }
                    else
                    {
                        labelIndicadorPercentagem.Text = "0%";
                    }
                }
            }
            BD.DesonectarBD();
        }

        private void CarregarPercentagensCircle()
        {
            ComunicaBD BD = new ComunicaBD();
            BD.ConectarBD();
            string NumeroObra = labelNumeroObra.Text.Trim();

            string percentagemEstrutura = string.Empty;
            string percentagemRevestimentos = string.Empty;
            string percentagemAprovação = string.Empty;
            string percentagemAlterações = string.Empty;
            string percentagemFabrico = string.Empty;
            string percentagemSoldadura = string.Empty;
            string percentagemMontagem = string.Empty;
            string percentagemDiversos = string.Empty;

            string query = @"
                            SELECT [Percentagem Estrutura], [Percentagem Revestimentos], [Percentagem Aprovação], 
                                   [Percentagem Alterações], [Percentagem Fabrico], [Percentagem Soldadura], 
                                   [Percentagem Montagem], [Percentagem Diversos]
                            FROM dbo.RealObras
                            WHERE [Numero da Obra] = @NumeroDaObra";

            using (SqlCommand cmd = new SqlCommand(query, BD.GetConnection()))
            {
                cmd.Parameters.AddWithValue("@NumeroDaObra", NumeroObra);

                using (SqlDataReader reader = cmd.ExecuteReader())
                {
                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            percentagemEstrutura = reader["Percentagem Estrutura"].ToString();
                            percentagemRevestimentos = reader["Percentagem Revestimentos"].ToString();
                            percentagemAprovação = reader["Percentagem Aprovação"].ToString();
                            percentagemAlterações = reader["Percentagem Alterações"].ToString();
                            percentagemFabrico = reader["Percentagem Fabrico"].ToString();
                            percentagemSoldadura = reader["Percentagem Soldadura"].ToString();
                            percentagemMontagem = reader["Percentagem Montagem"].ToString();
                            percentagemDiversos = reader["Percentagem Diversos"].ToString();
                        }

                        labelPercentagemEstrutura.Text = percentagemEstrutura;
                        labelPercentagemRevestimentos.Text = percentagemRevestimentos;
                        labelPercentagemAprov.Text = percentagemAprovação;
                        labelPercentagemAltera.Text = percentagemAlterações;
                        labelPercentagemFabrico.Text = percentagemFabrico;
                        labelPercentagemSoldadura.Text = percentagemSoldadura;
                        labelPercentagemMontagem.Text = percentagemMontagem;
                        labelPercentagemDiversos.Text = percentagemDiversos;

                    }
                    else
                    {

                        labelIndicadorPercentagem.Text = "0%";

                    }

                }
            }

            BD.DesonectarBD();
        }

        private void AtualizarIndicadorPercentagem()
        {
            string valorPercentagemStr = labelIndicadorPercentagem.Text.Replace("%", "").Trim();

            if (double.TryParse(valorPercentagemStr, out double percentagem))
            {
                percentagem = Math.Min(100, Math.Max(0, percentagem));

                guna2RadialGaugePercObra.Value = (int)percentagem;
            }
            else
            {
                MessageBox.Show("Valor de percentagem inválido.");
            }
        }

        public void CarregarGraficoPiePercentagem()
        {
            ComunicaBD BD = new ComunicaBD();
            BD.ConectarBD();
            string NumeroObra = labelNumeroObra.Text.Trim();

            string queryReal = @"
           SELECT [Percentagem Estrutura], [Percentagem Revestimentos], [Percentagem Aprovação],
           [Percentagem Alterações], [Percentagem Fabrico], [Percentagem Soldadura], 
           [Percentagem Montagem], [Percentagem Diversos]
            FROM dbo.RealObras
            WHERE [Numero da Obra] = @NumeroDaObra";

            chartCircle.Series["Percentagens"].Points.Clear();
            chartCircle.Series["Percentagens"].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Pie;
            chartCircle.Series["Percentagens"].IsValueShownAsLabel = true;
            chartCircle.Series["Percentagens"].IsValueShownAsLabel = true;

            chartCircle.Series["Percentagens"].Font = new Font("Arial", 13, FontStyle.Bold);

            using (SqlCommand cmd = new SqlCommand(queryReal, BD.GetConnection()))
            {
                cmd.Parameters.AddWithValue("@NumeroDaObra", NumeroObra);

                using (SqlDataReader reader = cmd.ExecuteReader())
                {
                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            double estruturaReal, revestimentosReal, aprovacaoReal, alteracoesReal, fabricoReal, soldaduraReal, montagemReal, diversosReal;

                            string estruturaRealStr = reader["Percentagem Estrutura"].ToString().Replace("%", "").Trim();
                            string revestimentosRealStr = reader["Percentagem Revestimentos"].ToString().Replace("%", "").Trim();
                            string aprovacaoRealStr = reader["Percentagem Aprovação"].ToString().Replace("%", "").Trim();
                            string alteracoesRealStr = reader["Percentagem Alterações"].ToString().Replace("%", "").Trim();
                            string fabricoRealStr = reader["Percentagem Fabrico"].ToString().Replace("%", "").Trim();
                            string soldaduraRealStr = reader["Percentagem Soldadura"].ToString().Replace("%", "").Trim();
                            string montagemRealStr = reader["Percentagem Montagem"].ToString().Replace("%", "").Trim();
                            string diversosRealStr = reader["Percentagem Diversos"].ToString().Replace("%", "").Trim();

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
                                int pontoIndex = chartCircle.Series["Percentagens"].Points.AddY(estruturaReal);
                                chartCircle.Series["Percentagens"].Points[pontoIndex].Color = Color.Red;
                                chartCircle.Series["Percentagens"].Points[pontoIndex].LabelForeColor = Color.Red;
                                chartCircle.Series["Percentagens"].Points[pontoIndex].BorderColor = Color.Red;
                            }
                            else
                            {
                                int pontoIndex = chartCircle.Series["Percentagens"].Points.AddY(0);
                                chartCircle.Series["Percentagens"].Points[pontoIndex].Color = Color.Red;
                                chartCircle.Series["Percentagens"].Points[pontoIndex].LabelForeColor = Color.Red;
                                chartCircle.Series["Percentagens"].Points[pontoIndex].BorderColor = Color.Red;
                            }

                            if (podeConverterRevestimentos)
                            {
                                int pontoIndex = chartCircle.Series["Percentagens"].Points.AddY(revestimentosReal);
                                chartCircle.Series["Percentagens"].Points[pontoIndex].Color = Color.FromArgb(97, 155, 243);
                                chartCircle.Series["Percentagens"].Points[pontoIndex].LabelForeColor = Color.FromArgb(97, 155, 243);
                                chartCircle.Series["Percentagens"].Points[pontoIndex].BorderColor = Color.FromArgb(97, 155, 243);
                            }
                            else
                            {
                                int pontoIndex = chartCircle.Series["Percentagens"].Points.AddY(0);
                                chartCircle.Series["Percentagens"].Points[pontoIndex].Color = Color.FromArgb(97, 155, 243);
                                chartCircle.Series["Percentagens"].Points[pontoIndex].LabelForeColor = Color.FromArgb(97, 155, 243);
                                chartCircle.Series["Percentagens"].Points[pontoIndex].BorderColor = Color.FromArgb(97, 155, 243);
                            }

                            if (podeConverterAprovacao)
                            {
                                int pontoIndex = chartCircle.Series["Percentagens"].Points.AddY(aprovacaoReal);
                                chartCircle.Series["Percentagens"].Points[pontoIndex].Color = Color.Orange;
                                chartCircle.Series["Percentagens"].Points[pontoIndex].LabelForeColor = Color.Orange;
                                chartCircle.Series["Percentagens"].Points[pontoIndex].BorderColor = Color.Orange;
                            }
                            else
                            {
                                int pontoIndex = chartCircle.Series["Percentagens"].Points.AddY(0);
                                chartCircle.Series["Percentagens"].Points[pontoIndex].Color = Color.Orange;
                                chartCircle.Series["Percentagens"].Points[pontoIndex].LabelForeColor = Color.Orange;
                                chartCircle.Series["Percentagens"].Points[pontoIndex].BorderColor = Color.Orange;
                            }

                            if (podeConverterAlteracoes)
                            {
                                int pontoIndex = chartCircle.Series["Percentagens"].Points.AddY(alteracoesReal);
                                chartCircle.Series["Percentagens"].Points[pontoIndex].Color = Color.FromArgb(139, 201, 77);
                                chartCircle.Series["Percentagens"].Points[pontoIndex].LabelForeColor = Color.FromArgb(139, 201, 77);
                                chartCircle.Series["Percentagens"].Points[pontoIndex].BorderColor = Color.FromArgb(139, 201, 77);
                            }
                            else
                            {
                                int pontoIndex = chartCircle.Series["Percentagens"].Points.AddY(0);
                                chartCircle.Series["Percentagens"].Points[pontoIndex].Color = Color.FromArgb(139, 201, 77);
                                chartCircle.Series["Percentagens"].Points[pontoIndex].LabelForeColor = Color.FromArgb(139, 201, 77);
                                chartCircle.Series["Percentagens"].Points[pontoIndex].BorderColor = Color.FromArgb(139, 201, 77);
                            }

                            if (podeConverterFabrico)
                            {
                                int pontoIndex = chartCircle.Series["Percentagens"].Points.AddY(fabricoReal);
                                chartCircle.Series["Percentagens"].Points[pontoIndex].Color = Color.FromArgb(255, 128, 255);
                                chartCircle.Series["Percentagens"].Points[pontoIndex].LabelForeColor = Color.FromArgb(255, 128, 255);
                                chartCircle.Series["Percentagens"].Points[pontoIndex].BorderColor = Color.FromArgb(255, 128, 255);
                            }
                            else
                            {
                                int pontoIndex = chartCircle.Series["Percentagens"].Points.AddY(0);
                                chartCircle.Series["Percentagens"].Points[pontoIndex].Color = Color.FromArgb(255, 128, 255);
                                chartCircle.Series["Percentagens"].Points[pontoIndex].LabelForeColor = Color.FromArgb(255, 128, 255);
                                chartCircle.Series["Percentagens"].Points[pontoIndex].BorderColor = Color.FromArgb(255, 128, 255);
                            }

                            if (podeConverterSoldadura)
                            {
                                int pontoIndex = chartCircle.Series["Percentagens"].Points.AddY(soldaduraReal);
                                chartCircle.Series["Percentagens"].Points[pontoIndex].Color = Color.DarkGreen;
                                chartCircle.Series["Percentagens"].Points[pontoIndex].LabelForeColor = Color.DarkGreen;
                                chartCircle.Series["Percentagens"].Points[pontoIndex].BorderColor = Color.DarkGreen;

                            }
                            else
                            {
                                int pontoIndex = chartCircle.Series["Percentagens"].Points.AddY(0);
                                chartCircle.Series["Percentagens"].Points[pontoIndex].Color = Color.DarkGreen;
                                chartCircle.Series["Percentagens"].Points[pontoIndex].LabelForeColor = Color.DarkGreen;
                                chartCircle.Series["Percentagens"].Points[pontoIndex].BorderColor = Color.DarkGreen;

                            }

                            if (podeConverterMontagem)
                            {
                                int pontoIndex = chartCircle.Series["Percentagens"].Points.AddY(montagemReal);
                                chartCircle.Series["Percentagens"].Points[pontoIndex].Color = Color.FromArgb(0, 192, 192); 
                                chartCircle.Series["Percentagens"].Points[pontoIndex].LabelForeColor = Color.FromArgb(0, 192, 192);
                                chartCircle.Series["Percentagens"].Points[pontoIndex].BorderColor = Color.FromArgb(0, 192, 192);

                            }
                            else
                            {
                                int pontoIndex = chartCircle.Series["Percentagens"].Points.AddY(0);
                                chartCircle.Series["Percentagens"].Points[pontoIndex].Color = Color.FromArgb(0, 192, 192);
                                chartCircle.Series["Percentagens"].Points[pontoIndex].LabelForeColor = Color.FromArgb(0, 192, 192);
                                chartCircle.Series["Percentagens"].Points[pontoIndex].BorderColor = Color.FromArgb(0, 192, 192);

                            }

                            if (podeConverterDiversos)
                            {
                                int pontoIndex = chartCircle.Series["Percentagens"].Points.AddY(diversosReal);
                                chartCircle.Series["Percentagens"].Points[pontoIndex].Color = Color.Gray;
                                chartCircle.Series["Percentagens"].Points[pontoIndex].LabelForeColor = Color.Gray;
                                chartCircle.Series["Percentagens"].Points[pontoIndex].BorderColor = Color.Gray;


                            }
                            else
                            {
                                int pontoIndex = chartCircle.Series["Percentagens"].Points.AddY(0);
                                chartCircle.Series["Percentagens"].Points[pontoIndex].Color = Color.Gray;
                                chartCircle.Series["Percentagens"].Points[pontoIndex].LabelForeColor = Color.Gray;
                                chartCircle.Series["Percentagens"].Points[pontoIndex].BorderColor = Color.Gray;

                            }

                        }
                    }
                }
            }

            BD.DesonectarBD();
        }

        private void CarregarPercentagensOrcamento()
        {
            ComunicaBD BD = new ComunicaBD();
            BD.ConectarBD();
            string NumeroObra = labelNumeroObra.Text.Trim();

            string KgEstrutura = string.Empty;
            string HorasEstrutura = string.Empty;
            string ValorEstrutura = string.Empty;
            string HorasRevestimento = string.Empty;
            string ValorRevestimento = string.Empty;

            string query = @"
                    SELECT [KG Estrutura], [Horas Estrutura], [Valor Estrutura], 
                           [Horas Revestimentos], [Valor Revestimentos]                           
                    FROM dbo.Orçamentação
                    WHERE [Numero da Obra] = @NumeroDaObra";

            using (SqlCommand cmd = new SqlCommand(query, BD.GetConnection()))
            {
                cmd.Parameters.AddWithValue("@NumeroDaObra", NumeroObra);

                using (SqlDataReader reader = cmd.ExecuteReader())
                {
                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            KgEstrutura = reader["KG Estrutura"].ToString();
                            HorasEstrutura = reader["Horas Estrutura"].ToString();
                            ValorEstrutura = reader["Valor Estrutura"].ToString();
                            HorasRevestimento = reader["Horas Revestimentos"].ToString();
                            ValorRevestimento = reader["Valor Revestimentos"].ToString();

                        }
                        ValorEstrutura = ValorEstrutura.Replace(".", ",");
                        ValorRevestimento = ValorRevestimento.Replace(".", ",");
                        HorasEstrutura = HorasEstrutura.Replace(",", ".");
                        HorasRevestimento = HorasRevestimento.Replace(",", ".");

                        labelKgEstruturaOrc.Text = KgEstrutura + " Kg";
                        labelHorasOrc.Text = HorasEstrutura + " h";
                        labelEuroOrc.Text = ValorEstrutura + " €";
                        labelHorasREVOrc.Text = HorasRevestimento + " h";
                        labelEuroOrcR.Text = ValorRevestimento + " €";
                    }

                }
            }

            BD.DesonectarBD();
        }

        private void CarregarPercentagensReal()
        {
            ComunicaBD BD = new ComunicaBD();
            BD.ConectarBD();
            string NumeroObra = labelNumeroObra.Text.Trim();

            string KgEstruturaReal = string.Empty;
            string HorasEstruturaReal = string.Empty;
            string ValorEstruturaReal = string.Empty;
            string HorasRevestimentoReal = string.Empty;
            string ValorRevestimentoReal = string.Empty;

            string query = @"
                    SELECT [KG Estrutura], [Horas Estrutura], [Valor Estrutura], 
                           [Horas Revestimentos], [Valor Revestimentos]                           
                    FROM dbo.RealObras
                    WHERE [Numero da Obra] = @NumeroDaObra";

            using (SqlCommand cmd = new SqlCommand(query, BD.GetConnection()))
            {
                cmd.Parameters.AddWithValue("@NumeroDaObra", NumeroObra);

                using (SqlDataReader reader = cmd.ExecuteReader())
                {
                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            KgEstruturaReal = reader["KG Estrutura"].ToString();
                            HorasEstruturaReal = reader["Horas Estrutura"].ToString();
                            ValorEstruturaReal = reader["Valor Estrutura"].ToString();
                            HorasRevestimentoReal = reader["Horas Revestimentos"].ToString();
                            ValorRevestimentoReal = reader["Valor Revestimentos"].ToString();

                        }
                        ValorEstruturaReal = ValorEstruturaReal.Replace(".", ",");
                        ValorRevestimentoReal = ValorRevestimentoReal.Replace(".", ",");
                        HorasEstruturaReal = HorasEstruturaReal.Replace(",", ".");
                        HorasRevestimentoReal = HorasRevestimentoReal.Replace(",", ".");

                        labelKgEstruturaReal.Text = KgEstruturaReal + " Kg";
                        labelHorasReall.Text = HorasEstruturaReal + " h";
                        labelEuroReal.Text = ValorEstruturaReal + " €";
                        labelHorasREVReal.Text = HorasRevestimentoReal + " h";
                        labelEuroReal2.Text = ValorRevestimentoReal + " €";
                    }

                }
            }

            BD.DesonectarBD();
        }
                
        private void CarregarValoresQuadro()
        {
            ComunicaBD BD = new ComunicaBD();
            BD.ConectarBD();
            string NumeroObra = labelNumeroObra.Text.Trim();

            string ValorEstrutura = string.Empty;
            string ValorRevestimentos = string.Empty;
            string ValorAprovação = string.Empty;
            string ValorAlterações = string.Empty;
            string ValorFabrico = string.Empty;
            string ValorSoldadura = string.Empty;
            string ValorMontagem = string.Empty;
            string ValorDiversos = string.Empty;

            string HorasEstrutura = string.Empty;
            string HorasRevestimentos = string.Empty;
            string HorasAprovação = string.Empty;
            string HorasAlterações = string.Empty;
            string HorasFabrico = string.Empty;
            string HorasSoldadura = string.Empty;
            string HorasMontagem = string.Empty;
            string HorasDiversos = string.Empty;

            string query = @"
                    SELECT [Valor Estrutura], [Valor Revestimentos], [Valor Aprovação], 
                           [Valor Alterações], [Valor Fabrico], [Valor Soldadura], 
                           [Valor Montagem], [Valor Diversos], 

                           [Horas Estrutura], [Horas Revestimentos], [Horas Aprovação], 
                           [Horas Alterações], [Horas Fabrico], [Horas Soldadura], 
                           [Horas Montagem], [Horas Diversos]
                    FROM dbo.RealObras
                    WHERE [Numero da Obra] = @NumeroDaObra";

            using (SqlCommand cmd = new SqlCommand(query, BD.GetConnection()))
            {
                cmd.Parameters.AddWithValue("@NumeroDaObra", NumeroObra);

                using (SqlDataReader reader = cmd.ExecuteReader())
                {
                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            //ValorEstrutura = reader["Valor Estrutura"].ToString() + " €";
                            //ValorRevestimentos = reader["Valor Revestimentos"].ToString() + " €";
                            //ValorAprovação = reader["Valor Aprovação"].ToString() + " €";
                            //ValorAlterações = reader["Valor Alterações"].ToString() + " €";
                            //ValorFabrico = reader["Valor Fabrico"].ToString() + " €";
                            //ValorSoldadura = reader["Valor Soldadura"].ToString() + " €";
                            //ValorMontagem = reader["Valor Montagem"].ToString() + " €";
                            //ValorDiversos = reader["Valor Diversos"].ToString() + " €";

                            decimal valorEstruturaDecimal = 0;
                            decimal valorRevestimentosDecimal = 0;
                            decimal valorAprovacaoDecimal = 0;
                            decimal valorAlteracoesDecimal = 0;
                            decimal valorFabricoDecimal = 0;
                            decimal valorSoldaduraDecimal = 0;
                            decimal valorMontagemDecimal = 0;
                            decimal valorDiversosDecimal = 0;

                            decimal.TryParse(reader["Valor Estrutura"].ToString(), out valorEstruturaDecimal);
                            decimal.TryParse(reader["Valor Revestimentos"].ToString(), out valorRevestimentosDecimal);
                            decimal.TryParse(reader["Valor Aprovação"].ToString(), out valorAprovacaoDecimal);
                            decimal.TryParse(reader["Valor Alterações"].ToString(), out valorAlteracoesDecimal);
                            decimal.TryParse(reader["Valor Fabrico"].ToString(), out valorFabricoDecimal);
                            decimal.TryParse(reader["Valor Soldadura"].ToString(), out valorSoldaduraDecimal);
                            decimal.TryParse(reader["Valor Montagem"].ToString(), out valorMontagemDecimal);
                            decimal.TryParse(reader["Valor Diversos"].ToString(), out valorDiversosDecimal);

                            ValorEstrutura = valorEstruturaDecimal.ToString("0.00") + " €";
                            ValorRevestimentos = valorRevestimentosDecimal.ToString("0.00") + " €";
                            ValorAprovação = valorAprovacaoDecimal.ToString("0.00") + " €";
                            ValorAlterações = valorAlteracoesDecimal.ToString("0.00") + " €";
                            ValorFabrico = valorFabricoDecimal.ToString("0.00") + " €";
                            ValorSoldadura = valorSoldaduraDecimal.ToString("0.00") + " €";
                            ValorMontagem = valorMontagemDecimal.ToString("0.00") + " €";
                            ValorDiversos = valorDiversosDecimal.ToString("0.00") + " €";

                            HorasEstrutura = reader["Horas Estrutura"].ToString() + " h";
                            HorasRevestimentos = reader["Horas Revestimentos"].ToString() + " h";
                            HorasAprovação = reader["Horas Aprovação"].ToString() + " h";
                            HorasAlterações = reader["Horas Alterações"].ToString() + " h";
                            HorasFabrico = reader["Horas Fabrico"].ToString() + " h";
                            HorasSoldadura = reader["Horas Soldadura"].ToString() + " h";
                            HorasMontagem = reader["Horas Montagem"].ToString() + " h";
                            HorasDiversos = reader["Horas Diversos"].ToString() + " h";
                        }

                        ValorEstrutura = ValorEstrutura.Replace(".", ",");
                        ValorRevestimentos = ValorRevestimentos.Replace(".", ",");
                        ValorAprovação = ValorAprovação.Replace(".", ",");
                        ValorAlterações = ValorAlterações.Replace(".", ",");
                        ValorFabrico = ValorFabrico.Replace(".", ",");
                        ValorSoldadura = ValorSoldadura.Replace(".", ",");
                        ValorMontagem = ValorMontagem.Replace(".", ",");
                        ValorDiversos = ValorDiversos.Replace(",", ".");

                        labelHorasEstrutura.Text = HorasEstrutura;
                        labelHorasRevestimentos.Text = HorasRevestimentos;
                        labelHorasAprov.Text = HorasAprovação;
                        labelHorasAltera.Text = HorasAlterações;
                        labelHorasFabrico.Text = HorasFabrico;
                        labelHorasSoldadura.Text = HorasSoldadura;
                        labelHorasMontagem.Text = HorasMontagem;
                        labelHorasDiversos.Text = HorasDiversos;

                        labelValorEstrutura.Text = ValorEstrutura;
                        labelValorRevestimentos.Text = ValorRevestimentos;
                        labelValorAprov.Text = ValorAprovação;
                        labelValorAltera.Text = ValorAlterações;
                        labelValorFabrico.Text = ValorFabrico;
                        labelValorSoldadura.Text = ValorSoldadura;
                        labelValorMontagem.Text = ValorMontagem;
                        labelValorDiversos.Text = ValorDiversos;

                    }

                }
            }

            BD.DesonectarBD();
        }

        private void ConfigurarDataGridView()
        {
            DataGridViewHorasPreparador.ReadOnly = true;

            if (DataGridViewHorasPreparador.Columns.Count == 0)
            {
                DataGridViewHorasPreparador.Columns.Add("Preparador", "Preparador");
                DataGridViewHorasPreparador.Columns.Add("TotalHoras", "Total de Horas");
                DataGridViewHorasPreparador.Columns.Add("Porcentagem", "Porcentagem");
            }
        }

        private void AtualizarTabelaHorasPreparador()
        {
            ConfigurarDataGridView();

            string NumeroObra = labelNumeroObra.Text.Trim();

            DataGridViewHorasPreparador.Rows.Clear();

            ComunicaBD BD = new ComunicaBD();

            try
            {
                BD.ConectarBD();

                List<PreparadorHoras> preparadores = ObterHorasTotaisPorPreparador(NumeroObra, BD);

                int totalHorasObra = preparadores.Sum(p => p.HorasTotais);

                foreach (var preparador in preparadores)
                {
                    double horas = Math.Round((double)preparador.HorasTotais / 60);

                    string horasFormatadas;
                    if (horas == Math.Floor(horas)) 
                    {
                        horasFormatadas = horas.ToString("0") + " h"; 
                    }
                    else
                    {
                        horasFormatadas = horas.ToString("0.00") + " h";
                    }

                    double porcentagem = (double)preparador.HorasTotais / totalHorasObra * 100;

                    DataGridViewHorasPreparador.Rows.Add(preparador.Nome, horasFormatadas, porcentagem.ToString("0.00") + "%");
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

        private List<PreparadorHoras> ObterHorasTotaisPorPreparador(string NumeroObra, ComunicaBD BD)
        {
            List<PreparadorHoras> preparadores = new List<PreparadorHoras>();

            string query = @"
                            SELECT Preparador, 
                                   SUM(DATEDIFF(MINUTE, '00:00:00', TRY_CAST([Qtd de Hora] AS TIME))) AS TotalHorasEmMinutos
                            FROM dbo.RegistoTempo
                            WHERE [Numero da Obra] = @NumeroObra
                            GROUP BY Preparador";

            using (SqlCommand cmd = new SqlCommand(query, BD.GetConnection()))
            {
                cmd.Parameters.AddWithValue("@NumeroObra", NumeroObra);

                using (SqlDataReader reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        PreparadorHoras preparador = new PreparadorHoras
                        {
                            Nome = reader.GetString(0), 
                            HorasTotais = reader.GetInt32(1) 
                        };

                        preparadores.Add(preparador);
                    }
                }
            }

            return preparadores;
        }

        public class PreparadorHoras
        {
            public string Nome { get; set; }
            public int HorasTotais { get; set; } 
        }

        private void CarregarHorasPorPreparador2()
        {
            ComunicaBD BD = new ComunicaBD();
            string NumeroObra = labelNumeroObra.Text.Trim();
            double totalHoras = 0;
            double totalHorasorcamentado = 0;

            try
            {
                BD.ConectarBD();

                string query = @"
                                 SELECT 
                                     [Total Horas]
                                 FROM 
                                     dbo.RealObras
                                 WHERE 
                                     [Numero da Obra] = @NumeroDaObra";

                string query2 = @"
                                 SELECT 
                                     [Total Horas]
                                 FROM 
                                     dbo.Orçamentação
                                 WHERE 
                                     [Numero da Obra] = @NumeroDaObra";

                using (SqlCommand cmd = new SqlCommand(query, BD.GetConnection()))
                {
                    cmd.Parameters.AddWithValue("@NumeroDaObra", NumeroObra);

                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        if (reader.HasRows)
                        {
                            while (reader.Read())
                            {
                                string totalHoras2 = reader.GetString(0);
                                totalHoras2 = totalHoras2.Replace(".", ",");
                                totalHoras = Convert.ToDouble(totalHoras2);

                                labelHorasReal.Text = totalHoras2 + " h";
                            }
                        }
                        else
                        {
                            MessageBox.Show("Nenhuma informação encontrada para a obra.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            return;
                        }
                    }

                    using (SqlCommand cmd2 = new SqlCommand(query2, BD.GetConnection()))
                    {
                        cmd2.Parameters.AddWithValue("@NumeroDaObra", NumeroObra);

                        using (SqlDataReader reader = cmd2.ExecuteReader())
                        {
                            if (reader.HasRows)
                            {
                                while (reader.Read())
                                {
                                    string totalHorasOrcamentado = reader.GetString(0);
                                    totalHorasOrcamentado = totalHorasOrcamentado.Replace(".", ",");
                                    totalHorasorcamentado = Convert.ToDouble(totalHorasOrcamentado);

                                    labelHorasOrcamentado.Text = totalHorasorcamentado + " h";
                                }
                            }
                            else
                            {
                                MessageBox.Show("Nenhuma informação encontrada para a obra.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                return;
                            }
                        }
                    }

                    //foreach (DataGridViewRow row in DataGridViewHorasPreparador.Rows)
                    //{
                    //    if (row.Cells["TotalHoras"].Value != null)
                    //    {
                    //        string totalHorasPreparador = row.Cells["TotalHoras"].Value.ToString();
                    //        totalHorasPreparador = totalHorasPreparador.Replace(":", ",");
                    //        double totalHorasPreparador2 = Convert.ToDouble(totalHorasPreparador);
                    //        double porcentagem = (totalHorasPreparador2 / totalHoras) * 100;
                    //        row.Cells["Porcentagem"].Value = porcentagem.ToString("F2") + "%";
                    //    }
                    //}

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erro ao carregar horas por preparador: " + ex.Message, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                BD.DesonectarBD();
            }

            
        }

        private void ConclusaoValores()
        {
            ComunicaBD BD = new ComunicaBD();
            BD.ConectarBD();
            string NumeroObra = labelNumeroObra.Text.Trim();

            string TotalHoras = string.Empty;
            string TotalValor = string.Empty;
            string DiasPreparados = string.Empty;

            string query = @"
             SELECT [Total Horas], [Total Valor], [Dias de Preparação]                          
             FROM dbo.ConclusaoObras
             WHERE [Numero da Obra] = @NumeroDaObra";

            using (SqlCommand cmd = new SqlCommand(query, BD.GetConnection()))
            {
                cmd.Parameters.AddWithValue("@NumeroDaObra", NumeroObra);

                using (SqlDataReader reader = cmd.ExecuteReader())
                {
                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            TotalHoras = reader["Total Horas"].ToString();
                            TotalValor = reader["Total Valor"].ToString();
                            DiasPreparados = reader["Dias de Preparação"].ToString();
                        }

                        if (string.IsNullOrEmpty(TotalHoras) || string.IsNullOrEmpty(TotalValor) || string.IsNullOrEmpty(DiasPreparados))
                        {
                            MessageBox.Show("Dados inválidos retornados do banco de dados.");
                            return;
                        }

                        TotalValor = TotalValor.Replace(".", ","); 
                        DiasPreparados = DiasPreparados.Replace("-", "").Replace(".", ",");

                        double diasPreparadosDouble = 0;
                        if (double.TryParse(DiasPreparados, out diasPreparadosDouble))
                        {
                            diasPreparadosDouble = Math.Round(diasPreparadosDouble);
                        }

                        
                        if (labelDiferncahoras.InvokeRequired)
                        {
                            labelDiferncahoras.Invoke(new Action(() =>
                            {
                                labelDiferncahoras.Text = TotalHoras + " h";
                            }));
                        }
                        else
                        {
                            labelDiferncahoras.Text = TotalHoras + " h";
                        }

                        if (labelDiferncaValor.InvokeRequired)
                        {
                            labelDiferncaValor.Invoke(new Action(() =>
                            {
                                labelDiferncaValor.Text = TotalValor + " €";
                            }));
                        }
                        else
                        {
                            labelDiferncaValor.Text = TotalValor + " €";
                        }                       
                        if (labelDias.InvokeRequired)
                        {
                            labelDias.Invoke(new Action(() =>
                            {
                             
                                labelDias.Text = ((int)diasPreparadosDouble).ToString() + " Dias";
                                
                            }));
                        }
                        else
                        {
                            labelDias.Text = diasPreparadosDouble.ToString() + " Dias";
                        }
                    }
                    else
                    {
                        MessageBox.Show("Nenhum dado encontrado para a obra " + NumeroObra);
                    }
                }
            }

            BD.DesonectarBD();
        }

        private void PercentagemHorasValor()
        {
            try
            {
                string horasorcamentadas = labelHorasOrc.Text.Trim();
                string horasreais = labelHorasReall.Text.Trim();

                horasorcamentadas = horasorcamentadas.Replace("h", "").Replace(".", ",");
                horasreais = horasreais.Replace("h", "").Replace(".", ",");


                if (double.TryParse(horasorcamentadas, out double horasOrcamento) &&
                    double.TryParse(horasreais, out double horasReais))
                {
                    if (horasOrcamento != 0)
                    {
                        double resultado = (horasReais / horasOrcamento) * 100;

                        LabelPercentagemH.Text = resultado.ToString("F1") + " %";
                    }
                    else
                    {
                        LabelPercentagemH.Text = "Valor de horas orçamentadas não pode ser zero.";
                    }
                }
                else
                {
                    LabelPercentagemH.Text = "Formato de hora inválido.";
                }
            }
            catch (Exception ex)
            {
                LabelPercentagemH.Text = "Erro ao processar as horas.";
            }


            try
            {
                string valororcamentadas = labelHorasREVOrc.Text.Trim();
                string valorreais = labelHorasREVReal.Text.Trim();

                valororcamentadas = valororcamentadas.Replace("h", "");
                valorreais = valorreais.Replace("h", "").Replace(".", ",");


                if (double.TryParse(valororcamentadas, out double valorOrcamento) &&
                    double.TryParse(valorreais, out double valorReais))
                {
                    if (valorOrcamento != 0)
                    {
                        double resultado = (valorReais / valorOrcamento) * 100;

                        labelPercentagemRev.Text = resultado.ToString("F1") + " %";
                    }
                    else
                    {
                        labelPercentagemRev.Text = "0 %";
                    }
                }
                else
                {
                    labelPercentagemRev.Text = "Formato de hora inválido.";
                }
            }
            catch (Exception ex)
            {
                LabelPercentagemH.Text = "Erro ao processar as horas.";
            }

        }

        private void ImagemObra()
        {
            string NumeroObra = labelNumeroObra.Text.Trim();

            string caminhoImagem = $".\\ImagensObra\\{NumeroObra}.png";

            if (File.Exists(caminhoImagem))
            {
                pictureBox1.Image = Image.FromFile(caminhoImagem);
            }
            else
            {
                MessageBox.Show("Imagem não encontrada para o número da obra informado.");
            }

        }

        private void ImagemObraTomb()
    {
        string NumeroObra = labelNumeroObra.Text.Trim();

        string doisPrimeirosDigitos = NumeroObra.Substring(0, 2);
        string ano = "20" + doisPrimeirosDigitos;
        string Caminhoinicio = @"\\marconi\COMPANY SHARED FOLDER\OFELIZ\OFM\2.AN\2.CM\DP\1 Obras";
        string CaminhoSecundario = @"\1.8 Projeto\1.8.2 Tekla\";
        string primeiraPasta = string.Empty; 
        string NomeFicheiro = "thumbnail.png";
        string Caminho = Caminhoinicio +  "\\" + ano + "\\" + NumeroObra + CaminhoSecundario;
                  
        if (Directory.Exists(Caminho))
        {
            string[] pastas = Directory.GetDirectories(Caminho);
            if (pastas.Length > 0)
            {
                primeiraPasta = Path.GetFileName(pastas[0]);                  
            }
            else
            {
                MessageBox.Show("Não há pastas no diretório especificado.");
            }
        }
        else
        {
            MessageBox.Show("O diretório não existe.");
        }

        string CaminhoCompleto = Path.Combine(Caminho, primeiraPasta, NomeFicheiro);
        pictureBox1.Image = Image.FromFile(CaminhoCompleto);

        }

        private void LabelFormatado()
        {
            labelEuroReal2.Text = FormatValor(labelEuroReal2.Text);
            labelDiferncaValor.Text = FormatValor(labelDiferncaValor.Text);
            labelEuroReal.Text = FormatValor(labelEuroReal.Text);

        }            

        private string FormatValor(string valor)
        {
            valor = valor.Replace("€", "").Replace("h", "").Trim();

            // Substituir vírgula por ponto para normalizar
            valor = valor.Replace(",", ".");

            if (decimal.TryParse(valor, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out decimal valorDecimal))
            {
                return valorDecimal.ToString("0.00", System.Globalization.CultureInfo.InvariantCulture) + " €";
            }
            else
            {
                return "Valor inválido";
            }
        }

        private void AtualizarImagemHoras()
         {
                string diferencaHoras = labelDiferncahoras.Text.Trim().Replace("h", "").Replace(".", ",");
                double diferencaHorasStr = 0;

                if (double.TryParse(diferencaHoras, out diferencaHorasStr))
                {
                    if (diferencaHorasStr <= 0)
                    {
                        pictureBoxhour.ImageLocation = @".\Imagens\goodhour.png";
                    }
                    else
                    {
                        pictureBoxhour.ImageLocation = @".\Imagens\badhour.png";
                    }
                }
                else
                {
                    MessageBox.Show("O valor de 'labelDiferncahoras' não é um número válido. Verifique o conteúdo do campo.", "Erro de Conversão", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

        private void AtualizarImagemHoras2()
         {
                string diferencaHorasorcamentadas = labelHorasOrc.Text.Trim().Replace("h", "").Replace(".", ",");
                string diferencaHorasreal = labelHorasReall.Text.Trim().Replace("h", "").Replace(".", ",");

                double horasOrcamentadas = 0;
                double horasReais = 0;

                if (double.TryParse(diferencaHorasorcamentadas, out horasOrcamentadas))
                {
                    if (double.TryParse(diferencaHorasreal, out horasReais))
                    {
                        double diferencaHoras = horasReais - horasOrcamentadas;

                        if (diferencaHoras <= 0)
                        {
                            pictureBox3.ImageLocation = @".\Imagens\goodhour.png";
                        }
                        else
                        {
                            pictureBox3.ImageLocation = @".\Imagens\badhour.png";
                        }
                    }
                    else
                    {
                        MessageBox.Show("O valor de 'labelHorasReall' não é um número válido. Verifique o conteúdo do campo.", "Erro de Conversão", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                else
                {
                    MessageBox.Show("O valor de 'labelHorasOrc' não é um número válido. Verifique o conteúdo do campo.", "Erro de Conversão", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

        private void AtualizarImagemHoras3()
         {
                string diferencaHorasorcamentadas = labelHorasREVOrc.Text.Trim().Replace("h", "").Replace(".", ",");
                string diferencaHorasreal = labelHorasREVReal.Text.Trim().Replace("h", "").Replace(".", ",");

                double horasOrcamentadas = 0;
                double horasReais = 0;

                if (double.TryParse(diferencaHorasorcamentadas, out horasOrcamentadas))
                {
                    if (double.TryParse(diferencaHorasreal, out horasReais))
                    {
                        double diferencaHoras = horasReais - horasOrcamentadas;

                        if (diferencaHoras <= 0)
                        {
                            pictureBox6.ImageLocation = @".\Imagens\goodhour.png";
                        }
                        else
                        {
                            pictureBox6.ImageLocation = @".\Imagens\badhour.png";
                        }
                    }
                    else
                    {
                        MessageBox.Show("O valor de 'labelHorasReall' não é um número válido. Verifique o conteúdo do campo.", "Erro de Conversão", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                else
                {
                    MessageBox.Show("O valor de 'labelHorasOrc' não é um número válido. Verifique o conteúdo do campo.", "Erro de Conversão", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
          }                          

        private void ButtonExcel_Click(object sender, EventArgs e)
        {
            ExcelsemData();
        }

        private void ExcelsemData()
        {
            string NumeroObra = labelNumeroObra.Text.Trim();

            string query = @"
                    SELECT Preparador, [Data da Tarefa], [Qtd de Hora], [Hora Inicial], [Hora Final], Prioridade, Tarefa
                    FROM dbo.RegistoTempo
                    WHERE [Numero da Obra] = @NumeroObra
                    ORDER BY [Data da Tarefa]";

            ComunicaBD comunicaBD = new ComunicaBD();
            ExcelExport excelExport = new ExcelExport();

            SqlCommand command = new SqlCommand(query, comunicaBD.GetConnection());

            command.Parameters.AddWithValue("@NumeroObra", NumeroObra);

            comunicaBD.ConectarBD();

            DataTable dataTable = comunicaBD.BuscarRegistros(command);
            string filePath = $@"C:\r\Registros_da_Obra_{NumeroObra}.xlsx";
            excelExport.ExportarParaExcel(dataTable, filePath, NumeroObra);
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

        private void guna2ImageButton1_Click(object sender, EventArgs e)
        {
            ExcelcomData();
        }

        private void ExcelcomData()
        {
            string NumeroObra = labelNumeroObra.Text.Trim();

            string query = @"
            SELECT Preparador, [Data da Tarefa], [Qtd de Hora], [Hora Inicial], [Hora Final], Prioridade, Tarefa
            FROM dbo.RegistoTempo
            WHERE [Numero da Obra] = @NumeroObra";

            if (DateTimePickerInicio.Value != DateTimePickerInicio.MinDate)
            {
                query += " AND [Data da Tarefa] >= @DataInicio";
            }

            if (DateTimePickerConclusao.Value != DateTimePickerConclusao.MinDate)
            {
                query += " AND [Data da Tarefa] <= @DataConclusao";
            }

            query += " ORDER BY [Data da Tarefa]";

            ComunicaBD comunicaBD = new ComunicaBD();
            ExcelExport excelExport = new ExcelExport();

            SqlCommand command = new SqlCommand(query, comunicaBD.GetConnection());

            command.Parameters.AddWithValue("@NumeroObra", NumeroObra);

            if (DateTimePickerInicio.Value != DateTimePickerInicio.MinDate)
            {
                command.Parameters.AddWithValue("@DataInicio", DateTimePickerInicio.Value.Date); 
            }

            if (DateTimePickerConclusao.Value != DateTimePickerConclusao.MinDate)
            {
                command.Parameters.AddWithValue("@DataConclusao", DateTimePickerConclusao.Value.Date); 
            }

            comunicaBD.ConectarBD();

            DataTable dataTable = comunicaBD.BuscarRegistros(command);
            string filePath = $@"C:\r\Registros_da_Obra_{NumeroObra}_com_Datas_especificas.xlsx";
            excelExport.ExportarParaExcel(dataTable, filePath, NumeroObra);

            comunicaBD.DesonectarBD();

            try
            {
                // Tenta abrir o arquivo Excel
                System.Diagnostics.Process.Start(filePath);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erro ao tentar abrir o arquivo Excel: " + ex.Message, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void GraficoTotalhoras()
        {
            chartTotalHoras.Series["Horas"].Points.Clear();

            foreach (DataGridViewRow row in DataGridViewHorasPreparador.Rows)
            {
                if (!row.IsNewRow)
                {
                    string nomePreparador = row.Cells[0].Value?.ToString();
                    string valorBruto = row.Cells[1].Value?.ToString();

                    if (!string.IsNullOrEmpty(valorBruto))
                    {
                        string valorLimpo = Regex.Replace(valorBruto, @"[^\d,.-]", "").Trim();

                        double totalHoras;
                        if (double.TryParse(valorLimpo, NumberStyles.Any, CultureInfo.CurrentCulture, out totalHoras))
                        {
                            chartTotalHoras.Series["Horas"].Points.AddXY(nomePreparador, totalHoras);
                        }
                        else
                        {
                            MessageBox.Show($"Erro ao converter valor: {valorBruto} (limpo: {valorLimpo})");
                        }
                    }
                }
            }
        }

        private void GraficoTotalPercentagem()
        {
            chartTotalPercentagem.Series["Percentagem"].Points.Clear();

               foreach (DataGridViewRow row in DataGridViewHorasPreparador.Rows)
               {
                if (!row.IsNewRow)
                 {
                    string nomePreparador = row.Cells[0].Value?.ToString();
                    string valorBruto = row.Cells[2].Value?.ToString();

                    if (!string.IsNullOrEmpty(valorBruto))
                    {
                        string valorLimpo = Regex.Replace(valorBruto, @"[^\d,.-]", "").Trim();

                        double totalHoras;
                        if (double.TryParse(valorLimpo, NumberStyles.Any, CultureInfo.CurrentCulture, out totalHoras))
                        {
                            chartTotalPercentagem.Series["Percentagem"].Points.AddXY(nomePreparador, totalHoras);
                        }
                        else
                        {
                            MessageBox.Show($"Erro ao converter valor: {valorBruto} (limpo: {valorLimpo})");
                        }
                    }
                 }
           }

        }
               

    }
    }


