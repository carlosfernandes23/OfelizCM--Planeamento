using LiveCharts;
using LiveCharts.Wpf;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Windows.Forms;
using System.Windows.Media;

namespace OfelizCM
{
    internal abstract class BaseConsulta
    {
        protected readonly ComunicaBD _conectarbd;
        protected BaseConsulta()
        {
            _conectarbd = new ComunicaBD(); 
        }
    }
    internal class Mostartabelas : BaseConsulta
    {
        public DataTable TabelaOrcamentacao(bool semanofecho = false)
        {
            try
            {
                _conectarbd.ConectarBD();
                string query = @"SELECT Id, [Ano de fecho], [Numero da Obra], [Nome da Obra], 
                                    [Preparador Responsavel], Tipologia, [KG Estrutura], 
                                    [Horas Estrutura], [Valor Estrutura], [KG/Euro Estrutura], 
                                    [Horas Revestimentos], [Valor Revestimentos], 
                                    [Total Horas], [Total Valor] 
                                    FROM dbo.Orçamentação";
                if (semanofecho)
                    query += " WHERE [Ano de fecho] IS NULL";

                DataTable dataTable = _conectarbd.Procurarbd(query);
                foreach (DataRow row in dataTable.Rows)
                {
                    for (int i = 0; i < dataTable.Columns.Count; i++)
                    {
                        if (row[i] != DBNull.Value && row[i] is string)
                            row[i] = ((string)row[i]).Trim();
                    }
                }
                return dataTable;
            }
            catch (Exception ex)
            {
                throw new Exception("Erro  Carregar Tabela Orçamentação.");
            }
            finally
            {
                _conectarbd.DesonectarBD();
            }
        }
        public DataTable TabelaReal(bool semanofecho = false)
        {
            try
            {
                _conectarbd.ConectarBD();
                string query = @"SELECT ID, [Ano de fecho], [Numero da Obra], Tipologia, 
                                    [KG Estrutura], [Horas Estrutura], [Valor Estrutura], [KG/Euro Estrutura], [Percentagem Estrutura], 
                                    [Horas Revestimentos], [Valor Revestimentos], [Percentagem Revestimentos], 
                                    [Horas Aprovação], [Valor Aprovação], [Percentagem Aprovação], 
                                    [Horas Alterações], [Valor Alterações], [Percentagem Alterações], 
                                    [Horas Fabrico], [Valor Fabrico], [Percentagem Fabrico], 
                                    [Horas Soldadura], [Valor Soldadura], [Percentagem Soldadura], 
                                    [Horas Montagem], [Valor Montagem], [Percentagem Montagem], 
                                    [Horas Diversos], [Valor Diversos], [Percentagem Diversos], [Comentario Diversos], 
                                    [Total Horas], [Total Valor] 
                                    FROM dbo.RealObras";
                if (semanofecho)
                    query += " WHERE [Ano de fecho] IS NULL";

                DataTable dataTable = _conectarbd.Procurarbd(query);
                foreach (DataRow row in dataTable.Rows)
                {
                    for (int i = 0; i < dataTable.Columns.Count; i++)
                    {
                        if (row[i] != DBNull.Value && row[i] is string)
                            row[i] = ((string)row[i]).Trim();
                    }
                }
                return dataTable;
            }
            catch (Exception ex)
            {
                throw new Exception("Erro ao Carregar Tabela Real.");
            }
            finally
            {
                _conectarbd.DesonectarBD();
            }
        }
        public DataTable TabelaConclusao(bool semanofecho = false)
        {
            try
            {
                _conectarbd.ConectarBD();
                string query = @"SELECT ID, [Ano de fecho], [Numero da Obra], Tipologia,
                                            [Total Horas], [Total Valor], [Percentagem Total], 
                                            [Dias de Preparação] 
                                            FROM dbo.ConclusaoObras";
                if (semanofecho)
                    query += " WHERE [Ano de fecho] IS NULL";

                DataTable dataTable = _conectarbd.Procurarbd(query);

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
                return dataTable;
            }
            catch (Exception ex)
            {
                throw new Exception("Erro ao  Carregar Tabela Conclusão.");
            }
            finally
            {
                _conectarbd.DesonectarBD();
            }
        }
        public DataTable TabelaOrçamentacaoTotal(bool semanofecho = false)
        {
            try
            {
                _conectarbd.ConectarBD();
                string query = @"SELECT ID, [Total KG Estrutura Orc], [Total Horas Estrutura Orc], 
                                            [Total Valor Estrutura Orc], [Total KG/Euro Estrutura Orc], [Total Horas Revestimentos Orc], 
                                            [Total Valor Revestimentos Orc], [Total Horas Orc], [Total Valor Orc]  
                                            FROM dbo.TotalObras";
                if (semanofecho)
                    query += " WHERE [Ano de fecho] IS NULL";

                DataTable dataTable = _conectarbd.Procurarbd(query);
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
                return dataTable;
            }
            catch (Exception ex)
            {
                throw new Exception("Erro ao  Carregar Tabela do total da Orçamentação.");
            }
            finally
            {
                _conectarbd.DesonectarBD();
            }
        }
        public DataTable TabelaRealTotal(bool semanofecho = false)
        {
            try
            {
                _conectarbd.ConectarBD();
                string query = @"SELECT ID, [Total KG Estrutura Real], [Total Horas Estrutura Real], [Total Valor Estrutura Real], [Total KG/Euro Estrutura Real], 
                                            [Percentagem Estrutura Real],[Total Horas Revestimentos Real], [Total Valor Revestimentos Real], [Percentagem Revestimentos Real], 
                                            [Total Horas Aprovacao Real] ,[Total Valor Aprovacao Real],[Percentagem Aprovacao Real], [Total Horas Alteracoes Real], 
                                            [Total Valor Alteracoes Real], [Percentagem Alteracoes Real], [Total Horas Fabrico Real], [Total Valor Fabrico Real], 
                                            [Percentagem Fabrico Real], [Total Horas Soldadura Real], [Total Valor Soldadura Real], [Percentagem Soldadura Real], 
                                            [Total Horas Montagem Real], [Total Valor Montagem Real],[Percentagem Montagem Real], [Total Horas Diversos Real], 
                                            [Total Valor Diversos Real], [Percentagem Diversos Real],  [Comentario Diversos Real], [Total Horas Real], [Total Valor Real]  
                                            FROM dbo.TotalObras";
                if (semanofecho)
                    query += " WHERE [Ano de fecho] IS NULL";

                DataTable dataTable = _conectarbd.Procurarbd(query);
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
                return dataTable;
            }
            catch (Exception ex)
            {
                throw new Exception("Erro ao  Carregar Tabela do total dos Valores Reais.");
            }
            finally
            {
                _conectarbd.DesonectarBD();
            }
        }
        public DataTable TabelaConclusaoTotal(bool semanofecho = false)
        {
            try
            {
                _conectarbd.ConectarBD();
                string query = @"SELECT ID, [Total Horas Concl], [Total Valor Concl], [Percentagem Total Concl], [Dias de Preparacao Concl] 
                               FROM dbo.TotalObras";
                if (semanofecho)
                    query += " WHERE [Ano de fecho] IS NULL";

                DataTable dataTable = _conectarbd.Procurarbd(query);
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
                return dataTable;
            }
            catch (Exception ex)
            {
                throw new Exception("Erro ao  Carregar Tabela do total da Conclusão.");
            }
            finally
            {
                _conectarbd.DesonectarBD();
            }
        }
    }
    internal class MostarGraficos : BaseConsulta
    {
        public SeriesCollection CarregarGraficoRedondo()
        {
            _conectarbd.ConectarBD();
            string query = @"SELECT [Percentagem Estrutura Real], [Percentagem Revestimentos Real], [Percentagem Aprovacao Real],
                                    [Percentagem Alteracoes Real], [Percentagem Fabrico Real], [Percentagem Soldadura Real], 
                                    [Percentagem Montagem Real], [Percentagem Diversos Real]
                                    FROM dbo.TotalObras";
             
            var series = new SeriesCollection();
            try
            {
                using (var cmd = new SqlCommand(query, _conectarbd.GetConnection()))
                using (var reader = cmd.ExecuteReader())
                {
                    if (reader.Read())
                    {
                        double ParsePercent(string col) =>
                            double.TryParse(reader[col].ToString().Replace("%", "").Trim(), out var val) ? val : 0;

                        var categorias = new Dictionary<string, (string Coluna, System.Windows.Media.Brush Cor)>
                            {
                                { "Estrutura", ("Percentagem Estrutura Real", new System.Windows.Media.SolidColorBrush(System.Windows.Media.Colors.Red)) },
                                { "Revestimentos", ("Percentagem Revestimentos Real", new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(97, 155, 243))) },
                                { "Aprovação", ("Percentagem Aprovacao Real", new System.Windows.Media.SolidColorBrush(System.Windows.Media.Colors.Orange)) },
                                { "Alterações", ("Percentagem Alteracoes Real", new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(139, 201, 77))) },
                                { "Fabrico", ("Percentagem Fabrico Real", new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(255, 128, 255))) },
                                { "Soldadura", ("Percentagem Soldadura Real", new System.Windows.Media.SolidColorBrush(System.Windows.Media.Colors.DarkGreen)) },
                                { "Montagem", ("Percentagem Montagem Real", new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(0, 192, 192))) },
                                { "Diversos", ("Percentagem Diversos Real", new System.Windows.Media.SolidColorBrush(System.Windows.Media.Colors.Gray)) }
                            };

                        foreach (var cat in categorias)
                        {
                            double valor = ParsePercent(cat.Value.Coluna);
                            var serie = new LiveCharts.Wpf.PieSeries
                            {
                                Title = cat.Key,
                                Values = new ChartValues<double> { valor },
                                DataLabels = true,
                                LabelPoint = chartPoint => Math.Round(chartPoint.Y, 0, MidpointRounding.AwayFromZero).ToString() + " %",
                                Fill = cat.Value.Cor,
                                FontSize = 12
                            };
                            series.Add(serie);
                        }
                    }
                }
                return series;
            }
            finally
            {
                _conectarbd.DesonectarBD();
            }
        }
        public SeriesCollection CarregarGraficoRedondoanodefecho(string Anofecho)
        {
            _conectarbd.ConectarBD();
            string query = @"SELECT [Percentagem Estrutura], [Percentagem Revestimentos], [Percentagem Aprovação],
                            [Percentagem Alterações], [Percentagem Fabrico], [Percentagem Soldadura], 
                            [Percentagem Montagem], [Percentagem Diversos]
                            FROM dbo.RealObras                        
                            WHERE [Ano de fecho] = @AnoFecho
                            ORDER BY ID ASC";

            var series = new SeriesCollection();
            try
            {
                using (var cmd = new SqlCommand(query, _conectarbd.GetConnection()))
                {
                    cmd.Parameters.AddWithValue("@AnoFecho", Anofecho);

                    using (var reader = cmd.ExecuteReader())
                    {
                        var totais = new Dictionary<string, double>
                            {
                                { "Percentagem Estrutura", 0 },
                                { "Percentagem Revestimentos", 0 },
                                { "Percentagem Aprovação", 0 },
                                { "Percentagem Alterações", 0 },
                                { "Percentagem Fabrico", 0 },
                                { "Percentagem Soldadura", 0 },
                                { "Percentagem Montagem", 0 },
                                { "Percentagem Diversos", 0 }
                            };

                        while (reader.Read())
                        {
                            foreach (var key in totais.Keys.ToList())
                            {
                                string raw = reader[key].ToString().Replace("%", "").Replace(".", ",").Trim();
                                if (double.TryParse(raw, out double val))
                                    totais[key] += val;
                            }
                        }

                        var categorias = new Dictionary<string, (string Coluna, System.Windows.Media.Brush Cor)>
                            {
                                { "Estrutura", ("Percentagem Estrutura", new System.Windows.Media.SolidColorBrush(System.Windows.Media.Colors.Red)) },
                                { "Revestimentos", ("Percentagem Revestimentos", new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(97, 155, 243))) },
                                { "Aprovação", ("Percentagem Aprovação", new System.Windows.Media.SolidColorBrush(System.Windows.Media.Colors.Orange)) },
                                { "Alterações", ("Percentagem Alterações", new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(139, 201, 77))) },
                                { "Fabrico", ("Percentagem Fabrico", new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(255, 128, 255))) },
                                { "Soldadura", ("Percentagem Soldadura", new System.Windows.Media.SolidColorBrush(System.Windows.Media.Colors.DarkGreen)) },
                                { "Montagem", ("Percentagem Montagem", new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(0, 192, 192))) },
                                { "Diversos", ("Percentagem Diversos", new System.Windows.Media.SolidColorBrush(System.Windows.Media.Colors.Gray)) }
                            };

                        double somaTotal = totais.Values.Sum();

                        foreach (var cat in categorias)
                        {
                            double valor = totais[cat.Value.Coluna];
                            double percentual = somaTotal > 0 ? valor / somaTotal * 100 : 0;

                            var serie = new LiveCharts.Wpf.PieSeries
                            {
                                Title = cat.Key,
                                Values = new ChartValues<double> { percentual },
                                DataLabels = true,
                                LabelPoint = chartPoint => Math.Round(chartPoint.Y, 0, MidpointRounding.AwayFromZero).ToString() + " %",
                                Fill = cat.Value.Cor,
                                FontSize = 12
                            };
                            series.Add(serie);
                        }
                    }
                }
                return series;
            }
            finally
            {
                _conectarbd.DesonectarBD();
            }
        }
        public SeriesCollection CarregarGraficoPiePercentagemComObra(string NumeroObra)
        {
            ComunicaBD BD = new ComunicaBD();
            BD.ConectarBD();

            string queryReal = @"SELECT [Percentagem Estrutura], [Percentagem Revestimentos], [Percentagem Aprovação],
                               [Percentagem Alterações], [Percentagem Fabrico], [Percentagem Soldadura], 
                               [Percentagem Montagem], [Percentagem Diversos]
                               FROM dbo.RealObras
                               WHERE [Numero da Obra] = @NumeroDaObra";

            var series = new LiveCharts.SeriesCollection();

            try
            {
                using (SqlCommand cmd = new SqlCommand(queryReal, BD.GetConnection()))
                {
                    cmd.Parameters.AddWithValue("@NumeroDaObra", NumeroObra);
                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            double ParsePercent(string col)
                            {
                                var val = reader[col]?.ToString().Replace("%", "").Trim();
                                return double.TryParse(val, out double result) ? result : 0;
                            }

                            var categorias = new Dictionary<string, (string Coluna, System.Windows.Media.Brush Cor)>
                            {
                                { "Estrutura", ("Percentagem Estrutura", new System.Windows.Media.SolidColorBrush(System.Windows.Media.Colors.Red)) },
                                { "Revestimentos", ("Percentagem Revestimentos", new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(97, 155, 243))) },
                                { "Aprovação", ("Percentagem Aprovação", new System.Windows.Media.SolidColorBrush(System.Windows.Media.Colors.Orange)) },
                                { "Alterações", ("Percentagem Alterações", new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(139, 201, 77))) },
                                { "Fabrico", ("Percentagem Fabrico", new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(255, 128, 255))) },
                                { "Soldadura", ("Percentagem Soldadura", new System.Windows.Media.SolidColorBrush(System.Windows.Media.Colors.DarkGreen)) },
                                { "Montagem", ("Percentagem Montagem", new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(0, 192, 192))) },
                                { "Diversos", ("Percentagem Diversos", new System.Windows.Media.SolidColorBrush(System.Windows.Media.Colors.Gray)) }
                            };

                            foreach (var cat in categorias)
                            {
                                double valor = ParsePercent(cat.Value.Coluna);

                                var pieSeries = new PieSeries
                                {
                                    Title = cat.Key,
                                    Values = new ChartValues<double> { valor },
                                    DataLabels = true,
                                    LabelPoint = point => Math.Round(point.Y).ToString() + " %",
                                    Fill = cat.Value.Cor,
                                    FontSize = 12
                                };

                                series.Add(pieSeries);
                            }
                        }
                    }
                }

                return series;
            }
            finally
            {
                BD.DesonectarBD();
            }
        }
        public SeriesCollection CalcularPercentagemTipologia(DataTable table)
        {
            if (table == null || table.Rows.Count == 0) return new SeriesCollection();

            double somaEstrutura = 0, somaRevestimentos = 0, somaAprovacao = 0;
            double somaAlteracoes = 0, somaFabrico = 0, somaSoldadura = 0;
            double somaMontagem = 0, somaDiversos = 0;
            int count = table.Rows.Count;

            foreach (DataRow row in table.Rows)
            {
                double ParseColumn(string columnName)
                {
                    var val = row[columnName]?.ToString().Replace("%", "").Trim();
                    return double.TryParse(val, out var d) ? d : 0;
                }

                somaEstrutura += ParseColumn("Percentagem Estrutura");
                somaRevestimentos += ParseColumn("Percentagem Revestimentos");
                somaAprovacao += ParseColumn("Percentagem Aprovação");
                somaAlteracoes += ParseColumn("Percentagem Alterações");
                somaFabrico += ParseColumn("Percentagem Fabrico");
                somaSoldadura += ParseColumn("Percentagem Soldadura");
                somaMontagem += ParseColumn("Percentagem Montagem");
                somaDiversos += ParseColumn("Percentagem Diversos");
            }

            return new SeriesCollection
            {
                new PieSeries
                {
                    Title = "Estrutura",
                    Values = new ChartValues<double> { somaEstrutura / count },
                    DataLabels = true,
                    FontSize = 12,
                    Fill = new SolidColorBrush(System.Windows.Media.Colors.Red),
                    LabelPoint = cp => Math.Round(cp.Y, 0, MidpointRounding.AwayFromZero).ToString() + " %"
                },
                new PieSeries
                {
                    Title = "Revestimentos",
                    Values = new ChartValues<double> { somaRevestimentos / count },
                    DataLabels = true,
                    Fill = new SolidColorBrush(System.Windows.Media.Color.FromRgb(97, 155, 243)),
                    LabelPoint = cp => Math.Round(cp.Y, 0, MidpointRounding.AwayFromZero).ToString() + " %"
                },
                new PieSeries
                {
                    Title = "Aprovação",
                    Values = new ChartValues<double> { somaAprovacao / count },
                    DataLabels = true,
                    FontSize = 12,
                    Fill = new SolidColorBrush(System.Windows.Media.Colors.Orange),
                    LabelPoint = cp => Math.Round(cp.Y, 0, MidpointRounding.AwayFromZero).ToString() + " %"
                },
                new PieSeries
                {
                    Title = "Alterações",
                    Values = new ChartValues<double> { somaAlteracoes / count },
                    DataLabels = true,
                    FontSize = 12,
                    Fill = new SolidColorBrush(System.Windows.Media.Color.FromRgb(139, 201, 77)),
                    LabelPoint = cp => Math.Round(cp.Y, 0, MidpointRounding.AwayFromZero).ToString() + " %"
                },
                new PieSeries
                {
                    Title = "Fabrico",
                    Values = new ChartValues<double> { somaFabrico / count },
                    DataLabels = true,
                    FontSize = 12,
                    Fill = new SolidColorBrush(System.Windows.Media.Color.FromRgb(255, 128, 255)),
                    LabelPoint = cp => Math.Round(cp.Y, 0, MidpointRounding.AwayFromZero).ToString() + " %"
                },
                new PieSeries
                {
                    Title = "Soldadura",
                    Values = new ChartValues<double> { somaSoldadura / count },
                    DataLabels = true,
                    FontSize = 12,
                    Fill = new SolidColorBrush(System.Windows.Media.Colors.DarkGreen),
                    LabelPoint = cp => Math.Round(cp.Y, 0, MidpointRounding.AwayFromZero).ToString() + " %"
                },
                new PieSeries
                {
                    Title = "Montagem",
                    Values = new ChartValues<double> { somaMontagem / count },
                    DataLabels = true,
                    FontSize = 12,
                    Fill = new SolidColorBrush(System.Windows.Media.Color.FromRgb(0, 192, 192)),
                    LabelPoint = cp => Math.Round(cp.Y, 0, MidpointRounding.AwayFromZero).ToString() + " %"
                },
                new PieSeries
                {
                    Title = "Diversos",
                    Values = new ChartValues<double> { somaDiversos / count },
                    DataLabels = true,
                    FontSize = 12,
                    Fill = new SolidColorBrush(System.Windows.Media.Colors.Gray),
                    LabelPoint = cp => Math.Round(cp.Y, 0, MidpointRounding.AwayFromZero).ToString() + " %"
                }
            };
        }
        public (SeriesCollection series, List<string> labels) CarregarPercentagemTodasObras(bool semanofecho = false)
        {
            _conectarbd.ConectarBD();

            string queryReal = "SELECT [Percentagem Total], [Numero da Obra] " +
                               "FROM dbo.ConclusaoObras";
                                if (semanofecho)
                                {
                                    queryReal += " WHERE [Ano de fecho] IS NULL";
                                }
                                queryReal += " ORDER BY ID ASC";

            var valores = new ChartValues<double>();
            var labels = new List<string>();

            try
            {
                using (SqlCommand cmd = new SqlCommand(queryReal, _conectarbd.GetConnection()))
                using (SqlDataReader reader = cmd.ExecuteReader())
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
                            valores.Add(totalValorPercentagem);
                            labels.Add(numeroDaObraReal);
                        }
                        else
                        {
                            valores.Add(0);
                            labels.Add(numeroDaObraReal);
                        }
                    }
                }

                var nomesObras = new Dictionary<string, string>();
                foreach (var numeroObra in labels)
                {
                    nomesObras[numeroObra] = ObterNomeObra(numeroObra);
                }

                var series = new SeriesCollection
                    {
                        new ColumnSeries
                        {
                            Title = "% Total",
                            Values = valores,
                            DataLabels = true,
                            Fill = Brushes.Yellow,
                            Stroke = Brushes.Black,
                            StrokeThickness = 0.5,
                            FontSize = 12,
                            LabelPoint = point =>
                            {
                              return $"{point.Y:N0} %";
                            }
                        }
                    };

                return (series, labels);
            }
            finally
            {
                _conectarbd.DesonectarBD();
            }
        }
        public (SeriesCollection series, List<string> labels) CarregarTotalTodasObras()
        {
            _conectarbd.ConectarBD();
            string query = @"
            SELECT [Total Horas Orc], [Total Horas Real]
            FROM dbo.TotalObras
            WHERE ID = @ID";

            var orcValues = new ChartValues<double>();
            var realValues = new ChartValues<double>();
            var labels = new List<string>();

            try
            {
                using (var cmd = new SqlCommand(query, _conectarbd.GetConnection()))
                {
                    cmd.Parameters.AddWithValue("@ID", "1");

                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            double totalOrc = 0, totalReal = 0;

                            string orcStr = reader["Total Horas Orc"].ToString().Replace("h", "").Trim();
                            string realStr = reader["Total Horas Real"].ToString().Replace("h", "").Trim();

                            double.TryParse(orcStr, out totalOrc);
                            double.TryParse(realStr, out totalReal);

                            orcValues.Add(totalOrc);
                            realValues.Add(totalReal);
                        }
                    }
                }

                var series = new SeriesCollection
            {
                new ColumnSeries
                {
                    Title = "Orçamentadas",
                    Values = orcValues,
                    DataLabels = true,
                    Fill = Brushes.LightBlue,
                    Stroke = Brushes.Black,
                    StrokeThickness = 0.5,
                    LabelPoint = point => point.Y + "h"
                },
                new ColumnSeries
                {
                    Title = "Real",
                    Values = realValues,
                    DataLabels = true,
                    Fill = Brushes.Orange,
                    Stroke = Brushes.Black,
                    StrokeThickness = 0.5,
                    LabelPoint = point => point.Y + "h"
                }
            };

                return (series, labels);
            }
            finally
            {
                _conectarbd.DesonectarBD();
            }
        }
        public (SeriesCollection series, List<string> labels) CarregarGraficoObrasHoras(bool semanofecho = false)
        {
            _conectarbd.ConectarBD();

            string queryOrc = @"SELECT [Total Horas], [Numero da Obra]
                                FROM dbo.Orçamentação";
                                if (semanofecho)
                                {
                                  queryOrc += " WHERE [Ano de fecho] IS NULL";
                                }
                                  queryOrc += " ORDER BY ID ASC";


            string queryReal = @"SELECT [Total Horas], [Numero da Obra]
                                FROM dbo.RealObras";
                                if (semanofecho)
                                {
                                  queryReal += " WHERE [Ano de fecho] IS NULL";
                                }
                                  queryReal += " ORDER BY ID ASC";

            var orcValues = new ChartValues<double>();
            var realValues = new ChartValues<double>();
            var labels = new List<string>();

            try
            {
                using (var cmd = new SqlCommand(queryOrc, _conectarbd.GetConnection()))
                using (var reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        string numeroDaObra = reader["Numero da Obra"].ToString();
                        string totalHorasStr = reader["Total Horas"].ToString().Replace("h", "").Trim();
                        totalHorasStr = totalHorasStr.Replace(".", ",");

                        if (!labels.Contains(numeroDaObra))
                            labels.Add(numeroDaObra);

                        if (!double.TryParse(totalHorasStr, out double totalHorasOrc))
                            totalHorasOrc = 0;

                        orcValues.Add(totalHorasOrc);
                    }
                }

                using (var cmd = new SqlCommand(queryReal, _conectarbd.GetConnection()))
                using (var reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        string totalHorasStr = reader["Total Horas"].ToString().Replace("h", "").Trim();
                        totalHorasStr = totalHorasStr.Replace(".", ",");

                        if (!double.TryParse(totalHorasStr, out double totalHorasReal))
                            totalHorasReal = 0;

                        realValues.Add(totalHorasReal);
                    }
                }
                var series = new SeriesCollection
                {
                        new ColumnSeries
                        {
                            Title = "Orçamentadas",
                            Values = orcValues,
                            DataLabels = true,
                            Fill = Brushes.LightBlue,
                            Stroke = Brushes.Black,
                            StrokeThickness = 0.5,
                            LabelPoint = point => point.Y + "h"
                        },
                        new ColumnSeries
                        {
                            Title = "Real",
                            Values = realValues,
                            DataLabels = true,
                            Fill = Brushes.Orange,
                            Stroke = Brushes.Black,
                            StrokeThickness = 0.5,
                            LabelPoint = point => point.Y + "h"
                        }
                    };

                return (series, labels);
            }
            finally
            {
                _conectarbd.DesonectarBD();
            }
        }
        public (SeriesCollection series, List<string> labels) CarregarGraficoObrasValor(bool semanofecho = false)
        {
            _conectarbd.ConectarBD();

            string queryOrc = @"SELECT [Total Valor], [Numero da Obra]
                                        FROM dbo.Orçamentação";
                                        if (semanofecho)
                                        {
                                         queryOrc += " WHERE [Ano de fecho] IS NULL";
                                        }
                                         queryOrc += " ORDER BY ID ASC";


            string queryReal = @"SELECT [Total Valor], [Numero da Obra]
                                        FROM dbo.RealObras";
                                        if (semanofecho)
                                        {
                                            queryReal += " WHERE [Ano de fecho] IS NULL";
                                        }
                                        queryReal += " ORDER BY ID ASC";

            var orcValues = new ChartValues<double>();
            var realValues = new ChartValues<double>();
            var labels = new List<string>();
            try
            {
                using (var cmd = new SqlCommand(queryOrc, _conectarbd.GetConnection()))
                using (var reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        string numeroDaObra = reader["Numero da Obra"].ToString();
                        string totalValorStr = reader["Total Valor"].ToString().Replace("€", "").Trim();
                        totalValorStr = totalValorStr.Replace(".", ",");

                        if (!labels.Contains(numeroDaObra))
                            labels.Add(numeroDaObra);

                        if (!double.TryParse(totalValorStr, out double totalValorOrc))
                            totalValorOrc = 0;

                        orcValues.Add(totalValorOrc);
                    }
                }
                using (var cmd = new SqlCommand(queryReal, _conectarbd.GetConnection()))
                using (var reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        string totalValorStr = reader["Total Valor"].ToString().Replace("€", "").Trim();
                        totalValorStr = totalValorStr.Replace(".", ",");

                        if (!double.TryParse(totalValorStr, out double totalValorReal))
                            totalValorReal = 0;

                        realValues.Add(totalValorReal);
                    }
                }
                var series = new SeriesCollection
                    {
                        new ColumnSeries
                        {
                            Title = "Orçamentadas",
                            Values = orcValues,
                            DataLabels = true,
                            Fill = Brushes.LightBlue,
                            Stroke = Brushes.Black,
                            StrokeThickness = 0.5,
                            LabelPoint = point => Math.Round(point.Y, 0) + "€",
                        },
                        new ColumnSeries
                        {
                            Title = "Real",
                            Values = realValues,
                            DataLabels = true,
                            Fill = Brushes.Orange,
                            Stroke = Brushes.Black,
                            StrokeThickness = 0.5,
                            LabelPoint = point => Math.Round(point.Y, 0) + "€",
                        }
                    };

                return (series, labels);
            }
            finally
            {
                _conectarbd.DesonectarBD();
            }
        }
        public (SeriesCollection series, List<string> labels) CarregarGraficoObrasValorTipologia(string campoFiltro, string valorFiltro , bool semanofecho = false)
        {
            _conectarbd.ConectarBD();
            string queryOrc = $@"SELECT [Total Valor], [Numero da Obra]
                                 FROM dbo.Orçamentação
                                 WHERE {campoFiltro} = @valor";
                                  if (semanofecho)
                                  {
                                    queryOrc += " AND [Ano de fecho] IS NULL";
                                  }
                                    queryOrc += " ORDER BY ID ASC";

            string queryReal = $@"SELECT [Total Valor], [Numero da Obra]
                                  FROM dbo.RealObras
                                  WHERE {campoFiltro} = @valor";
                                  if (semanofecho)
                                  {
                                    queryReal += " AND [Ano de fecho] IS NULL";
                                  }
                                    queryReal += " ORDER BY ID ASC";

            var orcValues = new ChartValues<double>();
            var realValues = new ChartValues<double>();
            var labels = new List<string>();
            try
            {
                using (var cmd = new SqlCommand(queryOrc, _conectarbd.GetConnection()))
                {
                    cmd.Parameters.AddWithValue("@valor", valorFiltro);
                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            string numeroDaObra = reader["Numero da Obra"].ToString();
                            string totalValorStr = reader["Total Valor"].ToString().Replace("€", "").Trim();
                            totalValorStr = totalValorStr.Replace(".", ",");

                            if (!labels.Contains(numeroDaObra))
                                labels.Add(numeroDaObra);

                            if (!double.TryParse(totalValorStr, out double totalValorOrc))
                                totalValorOrc = 0;

                            orcValues.Add(totalValorOrc);
                        }
                    }
                }
                using (var cmd = new SqlCommand(queryReal, _conectarbd.GetConnection()))
                {
                    cmd.Parameters.AddWithValue("@valor", valorFiltro);
                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            string totalValorStr = reader["Total Valor"].ToString().Replace("€", "").Trim();
                            totalValorStr = totalValorStr.Replace(".", ",");

                            if (!double.TryParse(totalValorStr, out double totalValorReal))
                                totalValorReal = 0;

                            realValues.Add(totalValorReal);
                        }
                    }
                }
                var series = new SeriesCollection
                        {
                            new ColumnSeries
                            {
                                Title = "Orçamentadas",
                                Values = orcValues,
                                DataLabels = true,
                                Fill = Brushes.LightBlue,
                                Stroke = Brushes.Black,
                                StrokeThickness = 0.5,
                                LabelPoint = point => Math.Round(point.Y, 0) + "€",
                            },
                            new ColumnSeries
                            {
                                Title = "Real",
                                Values = realValues,
                                DataLabels = true,
                                Fill = Brushes.Orange,
                                Stroke = Brushes.Black,
                                StrokeThickness = 0.5,
                                LabelPoint = point => Math.Round(point.Y, 0) + "€",
                            }
                        };

                return (series, labels);
            }
            finally
            {
                _conectarbd.DesonectarBD();
            }
        }
        public (SeriesCollection series, List<string> labels) CarregarGraficoObrasValorAno(string campoFiltro, string valorFiltro)
        {
            _conectarbd.ConectarBD();
            string queryOrc = $@"SELECT [Total Valor], [Numero da Obra]
                                 FROM dbo.Orçamentação
                                 WHERE {campoFiltro} = @valor ORDER BY ID ASC";

            string queryReal = $@"SELECT [Total Valor], [Numero da Obra]
                                  FROM dbo.RealObras
                                  WHERE {campoFiltro} = @valor ORDER BY ID ASC";

            var orcValues = new ChartValues<double>();
            var realValues = new ChartValues<double>();
            var labels = new List<string>();
            try
            {
                using (var cmd = new SqlCommand(queryOrc, _conectarbd.GetConnection()))
                {
                    cmd.Parameters.AddWithValue("@valor", valorFiltro);
                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            string numeroDaObra = reader["Numero da Obra"].ToString();
                            string totalValorStr = reader["Total Valor"].ToString().Replace("€", "").Trim();
                            totalValorStr = totalValorStr.Replace(".", ",");

                            if (!labels.Contains(numeroDaObra))
                                labels.Add(numeroDaObra);

                            if (!double.TryParse(totalValorStr, out double totalValorOrc))
                                totalValorOrc = 0;

                            orcValues.Add(totalValorOrc);
                        }
                    }
                }
                using (var cmd = new SqlCommand(queryReal, _conectarbd.GetConnection()))
                {
                    cmd.Parameters.AddWithValue("@valor", valorFiltro);
                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            string totalValorStr = reader["Total Valor"].ToString().Replace("€", "").Trim();
                            totalValorStr = totalValorStr.Replace(".", ",");

                            if (!double.TryParse(totalValorStr, out double totalValorReal))
                                totalValorReal = 0;

                            realValues.Add(totalValorReal);
                        }
                    }
                }
                var series = new SeriesCollection
                        {
                            new ColumnSeries
                            {
                                Title = "Orçamentadas",
                                Values = orcValues,
                                DataLabels = true,
                                Fill = Brushes.LightBlue,
                                Stroke = Brushes.Black,
                                StrokeThickness = 0.5,
                                LabelPoint = point => Math.Round(point.Y, 0) + "€",
                            },
                            new ColumnSeries
                            {
                                Title = "Real",
                                Values = realValues,
                                DataLabels = true,
                                Fill = Brushes.Orange,
                                Stroke = Brushes.Black,
                                StrokeThickness = 0.5,
                                LabelPoint = point => Math.Round(point.Y, 0) + "€",
                            }
                        };

                return (series, labels);
            }
            finally
            {
                _conectarbd.DesonectarBD();
            }
        }
        public (ChartValues<double> values, List<string> labels) CarregarGraficoObrasPercentagemAno(string anoFecho)
        {
            _conectarbd.ConectarBD();
            string queryReal = @"
            SELECT [Percentagem Total], [Numero da Obra]
            FROM dbo.ConclusaoObras
            WHERE [Ano de fecho] = @anoFecho";
            var values = new ChartValues<double>();
            var labels = new List<string>();
            try
            {
                using (var cmd = new SqlCommand(queryReal, _conectarbd.GetConnection()))
                {
                    cmd.Parameters.AddWithValue("@anoFecho", anoFecho);

                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            string numeroDaObra = reader["Numero da Obra"].ToString();
                            string valorStr = reader["Percentagem Total"].ToString().Replace("%", "").Trim();
                            valorStr = valorStr.Replace(".", ",");

                            if (!double.TryParse(valorStr, out double valor))
                                valor = 0;

                            values.Add(valor);

                            if (!labels.Contains(numeroDaObra))
                                labels.Add(numeroDaObra);
                        }
                    }
                }
                return (values, labels);
            }
            finally
            {
                _conectarbd.DesonectarBD();
            }
        }
        public (ChartValues<double> orcValues, ChartValues<double> realValues, List<string> labels) CarregarGraficoObrasHorasAno(string anoFecho)
        {
            _conectarbd.ConectarBD();
            string queryOrc = @"SELECT [Total Horas], [Numero da Obra]
                                FROM dbo.Orçamentação
                                WHERE [Ano de fecho] = @anoFecho";

            string queryReal = @"SELECT [Total Horas], [Numero da Obra]
                                 FROM dbo.RealObras
                                 WHERE [Ano de fecho] = @anoFecho";

            var orcValues = new ChartValues<double>();
            var realValues = new ChartValues<double>();
            var labels = new List<string>();

            try
            {
                using (var cmd = new SqlCommand(queryOrc, _conectarbd.GetConnection()))
                {
                    cmd.Parameters.AddWithValue("@anoFecho", anoFecho);

                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            string numeroDaObra = reader["Numero da Obra"].ToString();
                            string valorStr = reader["Total Horas"].ToString().Replace("h", "").Trim();
                            valorStr = valorStr.Replace(".", ",");

                            if (!double.TryParse(valorStr, out double valor))
                                valor = 0;

                            orcValues.Add(valor);
                            if (!labels.Contains(numeroDaObra))
                                labels.Add(numeroDaObra);
                        }
                    }
                }
                using (var cmd = new SqlCommand(queryReal, _conectarbd.GetConnection()))
                {
                    cmd.Parameters.AddWithValue("@anoFecho", anoFecho);

                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            string valorStr = reader["Total Horas"].ToString().Replace("h", "").Trim();
                            valorStr = valorStr.Replace(".", ",");

                            if (!double.TryParse(valorStr, out double valor))
                                valor = 0;

                            realValues.Add(valor);
                        }
                    }
                }
                return (orcValues, realValues, labels);
            }
            finally
            {
                _conectarbd.DesonectarBD();
            }
        }
        public (ChartValues<double> values, List<string> labels) CarregarGraficoObrasPercentagemTipologia(string tipologia, bool semanofecho = false)
        {
            _conectarbd.ConectarBD();
            string queryReal = @"SELECT [Percentagem Total], [Numero da Obra]
                                FROM dbo.ConclusaoObras
                                WHERE Tipologia = @tipologia";
                                if (semanofecho)
                                {
                                    queryReal += " AND [Ano de fecho] IS NULL";
                                }
            var values = new ChartValues<double>();
            var labels = new List<string>();
            try
            {
                using (var cmd = new SqlCommand(queryReal, _conectarbd.GetConnection()))
                {
                    cmd.Parameters.AddWithValue("@tipologia", tipologia);

                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            string numeroDaObra = reader["Numero da Obra"].ToString();
                            string valorStr = reader["Percentagem Total"].ToString().Replace("%", "").Trim();
                            valorStr = valorStr.Replace(".", ",");

                            if (!double.TryParse(valorStr, out double valor))
                                valor = 0;

                            values.Add(valor);
                            if (!labels.Contains(numeroDaObra))
                                labels.Add(numeroDaObra);
                        }
                    }
                }
                return (values, labels);
            }
            finally
            {
                _conectarbd.DesonectarBD();
            }
        }
        public (SeriesCollection series, List<string> labels) CarregarGraficoObrasHorasTipologia(string tipologia, bool semanofecho = false)
        {
            _conectarbd.ConectarBD();
            string queryOrc = @"SELECT [Total Horas], [Numero da Obra]
                                FROM dbo.Orçamentação
                                WHERE Tipologia = @Tipologia";
                                if (semanofecho) 
                                {
                                  queryOrc += " AND [Ano de fecho] IS NULL";
                                }
                                  queryOrc += " ORDER BY ID ASC";

            string queryReal = @"SELECT [Total Horas], [Numero da Obra]
                                FROM dbo.RealObras
                                WHERE Tipologia = @Tipologia";
                                if (semanofecho) 
                                {
                                    queryReal += " AND [Ano de fecho] IS NULL";
                                }
                                queryReal += " ORDER BY ID ASC";

            var orcValues = new ChartValues<double>();
            var realValues = new ChartValues<double>();
            var labels = new List<string>();
            try
            {
                using (var cmd = new SqlCommand(queryOrc, _conectarbd.GetConnection()))
                {
                    cmd.Parameters.AddWithValue("@Tipologia", tipologia);
                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            string numeroDaObra = reader["Numero da Obra"].ToString();
                            string totalHorasStr = reader["Total Horas"].ToString().Replace("h", "").Trim();
                            totalHorasStr = totalHorasStr.Replace(".", ",");

                            if (!double.TryParse(totalHorasStr, out double totalHorasOrc))
                                totalHorasOrc = 0;

                            orcValues.Add(totalHorasOrc);
                            if (!labels.Contains(numeroDaObra))
                                labels.Add(numeroDaObra);
                        }
                    }
                }
                using (var cmd = new SqlCommand(queryReal, _conectarbd.GetConnection()))
                {
                    cmd.Parameters.AddWithValue("@Tipologia", tipologia);
                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            string totalHorasStr = reader["Total Horas"].ToString().Replace("h", "").Trim();
                            totalHorasStr = totalHorasStr.Replace(".", ",");

                            if (!double.TryParse(totalHorasStr, out double totalHorasReal))
                                totalHorasReal = 0;

                            realValues.Add(totalHorasReal);
                        }
                    }
                }
                var series = new SeriesCollection
                    {
                        new ColumnSeries
                        {
                            Title = "Orçamentadas",
                            Values = orcValues,
                            DataLabels = true,
                            Fill = Brushes.LightBlue,
                            Stroke = Brushes.Black,
                            StrokeThickness = 0.5,
                            LabelPoint = point => point.Y + "h"
                        },
                        new ColumnSeries
                        {
                            Title = "Real",
                            Values = realValues,
                            DataLabels = true,
                            Fill = Brushes.Orange,
                            Stroke = Brushes.Black,
                            StrokeThickness = 0.5,
                            LabelPoint = point => point.Y + "h"
                        }
                    };

                return (series, labels);
            }
            finally
            {
                _conectarbd.DesonectarBD();
            }
        }
        public (SeriesCollection series, List<string> labels) CarregarGraficoHorasTotais(DataGridView dgvOrcamento, DataGridView dgvReal)
        {
            double totalHorasOrcamento = 0;
            double totalHorasReal = 0;
            foreach (DataGridViewRow row in dgvOrcamento.Rows)
            {
                if (row.Cells["Total Horas"].Value != null)
                {
                    string totalHorasOrcStr = row.Cells["Total Horas"].Value.ToString()
                        .Replace("h", "").Trim().Replace(".", ",");
                    if (!double.TryParse(totalHorasOrcStr, out double totalHorasOrc))
                        totalHorasOrc = 0;
                    totalHorasOrcamento += totalHorasOrc;
                }
            }
            foreach (DataGridViewRow row in dgvReal.Rows)
            {
                if (row.Cells["Total Horas"].Value != null)
                {
                    string totalHorasRealStr = row.Cells["Total Horas"].Value.ToString()
                        .Replace("h", "").Trim().Replace(".", ",");
                    if (!double.TryParse(totalHorasRealStr, out double totalHorasRealAux))
                        totalHorasRealAux = 0;
                    totalHorasReal += totalHorasRealAux;
                }
            }
            var series = new SeriesCollection
                        {
                            new ColumnSeries
                            {
                                Title = "Orçamentadas",
                                Values = new ChartValues<double> { totalHorasOrcamento },
                                DataLabels = true,
                                LabelPoint = point => point.Y + "h",
                                Fill = Brushes.LightBlue
                            },
                            new ColumnSeries
                            {
                                Title = "Real",
                                Values = new ChartValues<double> { totalHorasReal },
                                DataLabels = true,
                                LabelPoint = point => point.Y + "h",
                                Fill = Brushes.Orange
                            }
                        };

            var labels = new List<string> { "Total Horas" };

            return (series, labels);
        }
        public string ObterNomeObra(string numeroObra)
        {
            try
            {
                _conectarbd.ConectarBD();
                string query = @"SELECT [Nome da Obra] 
                         FROM dbo.[Orçamentação] 
                         WHERE [Numero da Obra] = @numeroObra";

                using (SqlCommand cmd = new SqlCommand(query, _conectarbd.GetConnection()))
                {
                    cmd.Parameters.AddWithValue("@numeroObra", numeroObra);
                    object result = cmd.ExecuteScalar();
                    return result?.ToString() ?? "";
                }
            }
            finally
            {
                _conectarbd.DesonectarBD();
            }
        }
    }
    internal class Carregarhorasgrafico : BaseConsulta
    {
        private readonly ComunicaBD _bd;
        public Carregarhorasgrafico(ComunicaBD bd)
        {
            _bd = bd;
        }
        public class PreparadorHoras
        {
            public string Nome { get; set; }
            public int HorasTotais { get; set; }
            public double Percentagem { get; set; }
        }
        public List<PreparadorHoras> ObterHorasPorPreparador(string numeroObra)
        {
            var preparadores = new List<PreparadorHoras>();
            string query = @"SELECT Preparador, 
                                   SUM(DATEDIFF(MINUTE, '00:00:00', TRY_CAST([Qtd de Hora] AS TIME))) AS TotalHorasEmMinutos
                            FROM dbo.RegistoTempo
                            WHERE [Numero da Obra] = @NumeroObra
                            GROUP BY Preparador";

            var connection = _bd.GetConnection();
            if (connection.State != System.Data.ConnectionState.Open)
                connection.Open();
            using (SqlCommand cmd = new SqlCommand(query, connection))
            {
                cmd.Parameters.AddWithValue("@NumeroObra", numeroObra);

                using (SqlDataReader reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        preparadores.Add(new PreparadorHoras
                        {
                            Nome = reader.GetString(0),
                            HorasTotais = reader.GetInt32(1)
                        });
                    }
                }
            }
            int totalHorasObra = preparadores.Sum(p => p.HorasTotais);
            foreach (var p in preparadores)
            {
                p.Percentagem = totalHorasObra == 0 ? 0 : (double)p.HorasTotais / totalHorasObra * 100;
            }
            return preparadores;
        }
    }
}
