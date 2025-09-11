using OfelizCM;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;

public class ExcelExport
{
    public void ExportarParaExcel(DataTable dataTable, string filePath, string numeroObra)
    {
        if (File.Exists(filePath))
        {
            File.Delete(filePath);
        }

        using (var package = new ExcelPackage())
        {
            var worksheet = package.Workbook.Worksheets.Add($"{numeroObra}");

            for (int i = 0; i < dataTable.Columns.Count; i++)
            {
                var cell = worksheet.Cells[2, i + 1];
                cell.Value = dataTable.Columns[i].ColumnName;

                cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                cell.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(155, 194, 230));
                cell.Style.Font.Bold = true;
                cell.Style.Font.Color.SetColor(Color.Black);
                cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                cell.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                cell.Style.Font.Size = 12;
                cell.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                worksheet.Row(2).Height = 30;

                switch (i)
                {
                    case 0: // Preparador
                        worksheet.Column(i + 1).Width = 19;
                        break;
                    case 1: // Data da Tarefa
                        worksheet.Column(i + 1).Width = 15;
                        break;
                    case 2: // Qtd de Hora
                        worksheet.Column(i + 1).Width = 13;
                        break;
                    case 3: // Hora Inicial
                        worksheet.Column(i + 1).Width = 13;
                        break;
                    case 4: // Hora Final
                        worksheet.Column(i + 1).Width = 13;
                        break;
                    case 5: // Prioridade
                        worksheet.Column(i + 1).Width = 40;
                        break;
                    case 6: // Tarefa
                        worksheet.Column(i + 1).Width = 100;
                        break;
                }
            }

            for (int row = 0; row < dataTable.Rows.Count; row++)
            {
                for (int col = 0; col < dataTable.Columns.Count; col++)
                {
                    var cell = worksheet.Cells[row + 3, col + 1];
                    cell.Value = dataTable.Rows[row][col].ToString();
                    cell.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                    cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    cell.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                }
            }

            FileInfo fi = new FileInfo(filePath);
            package.SaveAs(fi);
        }
    }

    public void ExportarParaExcelTodos(DataTable dataTable, string filePath)
    {
        if (File.Exists(filePath))
        {
            File.Delete(filePath);
        }

        using (var package = new ExcelPackage())
        {
            var worksheet = package.Workbook.Worksheets.Add($"Registos");

            for (int i = 0; i < dataTable.Columns.Count; i++)
            {
                var cell = worksheet.Cells[2, i + 1];
                cell.Value = dataTable.Columns[i].ColumnName;

                cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                cell.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(155, 194, 230));
                cell.Style.Font.Bold = true;
                cell.Style.Font.Color.SetColor(Color.Black);
                cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                cell.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                cell.Style.Font.Size = 12;
                cell.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                worksheet.Row(2).Height = 30;

                switch (i)
                {
                    case 0: // Numero Obra
                        worksheet.Column(i + 1).Width = 18;
                        break;
                    case 1: // Preparador
                        worksheet.Column(i + 1).Width = 19;
                        break;
                    case 2: // Data da Tarefa
                        worksheet.Column(i + 1).Width = 15;
                        break;
                    case 3: // Qtd de Hora
                        worksheet.Column(i + 1).Width = 13;
                        break;
                    case 4: // Hora Inicial
                        worksheet.Column(i + 1).Width = 13;
                        break;
                    case 5: // Hora Final
                        worksheet.Column(i + 1).Width = 13;
                        break;
                    case 6: // Prioridade
                        worksheet.Column(i + 1).Width = 40;
                        break;
                    case 7: // Tarefa
                        worksheet.Column(i + 1).Width = 100;
                        break;
                }
            }

            for (int row = 0; row < dataTable.Rows.Count; row++)
            {
                for (int col = 0; col < dataTable.Columns.Count; col++)
                {
                    var cell = worksheet.Cells[row + 3, col + 1];
                    cell.Value = dataTable.Rows[row][col].ToString();
                    cell.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                    cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    cell.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                }
            }

            FileInfo fi = new FileInfo(filePath);
            package.SaveAs(fi);
        }
    }

    public void ExportarParaExcelTabela(DataTable dataTable, ExcelWorksheet worksheet, int startColumn)
    {
        for (int i = 0; i < dataTable.Columns.Count; i++)
        {
            var cell = worksheet.Cells[2, startColumn + i];
            cell.Value = dataTable.Columns[i].ColumnName;

            if (startColumn + i >= 2 && startColumn + i <= 14) // Colunas 2 a 14
            {
                cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                cell.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(155, 194, 230)); // Cor azul
            }
            else if (startColumn + i >= 14 && startColumn + i <= 43) // Colunas 14 a 30
            {
                cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                cell.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(244, 176, 132)); // Cor laranja
            }
            else if (startColumn + i >= 43 && startColumn + i <= 47) // Colunas 31 a 34
            {
                cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                cell.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 192, 0)); // Cor amarela
            }

            cell.Style.Font.Bold = true;
            cell.Style.Font.Color.SetColor(Color.Black);
            cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            cell.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            cell.Style.Font.Size = 11;
            cell.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);

            worksheet.Row(2).Height = 30;

            switch (startColumn + i)
            {
                case 1:
                    worksheet.Column(startColumn + i).Width = 15;
                    break;
                case 2:
                    worksheet.Column(startColumn + i).Width = 15;
                    break;
                case 3:
                    worksheet.Column(startColumn + i).Width = 20;
                    break;
                case 4:
                    worksheet.Column(startColumn + i).Width = 60;
                    break;
                case 5:
                    worksheet.Column(startColumn + i).Width = 22;
                    break;
                case 6:
                    worksheet.Column(startColumn + i).Width = 20;
                    break;
                default: // Colunas restantes
                    worksheet.Column(startColumn + i).Width = 20;
                    break;
            }
        }

        for (int row = 0; row < dataTable.Rows.Count; row++)
        {
            for (int col = 0; col < dataTable.Columns.Count; col++)
            {
                var cell = worksheet.Cells[row + 3, startColumn + col];

                if (dataTable.Columns[col].ColumnName == "Numero da Obra" ||
                    dataTable.Columns[col].ColumnName == "Nome da Obra" ||
                    dataTable.Columns[col].ColumnName == "Preparador Responsavel" ||
                    dataTable.Columns[col].ColumnName == "Tipologia")
                {
                    cell.Value = dataTable.Rows[row][col].ToString().Trim();
                }
                else
                {
                    cell.Value = dataTable.Rows[row][col].ToString();
                }

                cell.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                cell.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            }
        }
    }


    //public void ExportarParaExcelContabelizar(DataTable dataTable, string filePath)
    //{
    //    using (var package = new ExcelPackage(new FileInfo(filePath)))
    //    {
    //        var worksheet = package.Workbook.Worksheets.FirstOrDefault(ws => ws.Name == "Registos");
    //        if (worksheet == null)
    //        {
    //            worksheet = package.Workbook.Worksheets.Add("Registos");
    //        }

    //        var borderColor = Color.FromArgb(155, 194, 230); // Cor da borda azul
    //        var thinBorderStyle = ExcelBorderStyle.Thin; // Estilo da borda fina
    //        var oddRowColor = Color.FromArgb(221, 235, 247); // Cor de fundo para linhas ímpares

    //        for (int i = 0; i < dataTable.Columns.Count; i++)
    //        {
    //            var cell = worksheet.Cells[2, i + 1];

    //            switch (i)
    //            {
    //                case 0:
    //                    cell.Value = "Data";
    //                    break;
    //                case 1:
    //                    cell.Value = "Obra";
    //                    break;
    //                case 3:
    //                    cell.Value = "Codigo";
    //                    break;
    //                case 6:
    //                    cell.Value = "Observações";
    //                    break;
    //                case 7:
    //                    cell.Value = "Designação";
    //                    break;
    //                default:
    //                    cell.Value = dataTable.Columns[i].ColumnName;
    //                    break;
    //            }

    //            cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
    //            cell.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(155, 194, 230));
    //            cell.Style.Font.Bold = true;
    //            cell.Style.Font.Color.SetColor(Color.Black);
    //            cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
    //            cell.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
    //            cell.Style.Font.Size = 12;
    //            worksheet.Row(2).Height = 30;

    //            switch (i)
    //            {
    //                case 0: // Data
    //                    worksheet.Column(i + 1).Width = 15;
    //                    break;
    //                case 1: // Obra
    //                    worksheet.Column(i + 1).Width = 15;
    //                    break;
    //                case 2: // Preparador
    //                    worksheet.Column(i + 1).Width = 20;
    //                    break;
    //                case 3: // Codigo
    //                    worksheet.Column(i + 1).Width = 18;
    //                    break;
    //                case 4: // Hora Inicial
    //                    worksheet.Column(i + 1).Width = 13;
    //                    break;
    //                case 5: // Hora Final
    //                    worksheet.Column(i + 1).Width = 13;
    //                    break;
    //                case 6: // Observações
    //                    worksheet.Column(i + 1).Width = 20;
    //                    break;
    //                case 7: // Designação
    //                    worksheet.Column(i + 1).Width = 45;
    //                    break;
    //                case 8: // Qtd de Hora
    //                    worksheet.Column(i + 1).Width = 14;
    //                    break;
    //            }
    //        }

    //        for (int row = 0; row < dataTable.Rows.Count; row++)
    //        {
    //            for (int col = 0; col < dataTable.Columns.Count; col++)
    //            {
    //                var cell = worksheet.Cells[row + 3, col + 1];
    //                string value = dataTable.Rows[row][col].ToString();

    //                if (col == 6 && value == "0")
    //                {
    //                    cell.Value = "";
    //                }
    //                else
    //                {
    //                    cell.Value = value;
    //                }

    //                cell.Style.Border.Top.Style = thinBorderStyle;
    //                cell.Style.Border.Left.Style = thinBorderStyle;
    //                cell.Style.Border.Right.Style = thinBorderStyle;
    //                cell.Style.Border.Bottom.Style = thinBorderStyle;

    //                cell.Style.Border.Top.Color.SetColor(borderColor);
    //                cell.Style.Border.Left.Color.SetColor(borderColor);
    //                cell.Style.Border.Right.Color.SetColor(borderColor);
    //                cell.Style.Border.Bottom.Color.SetColor(borderColor);

    //                cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
    //                cell.Style.VerticalAlignment = ExcelVerticalAlignment.Center;

    //                if ((row + 1) % 2 != 0)
    //                {
    //                    cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
    //                    cell.Style.Fill.BackgroundColor.SetColor(oddRowColor);
    //                }
    //            }
    //        }

    //        var headerRange = worksheet.Cells[2, 1, 2, dataTable.Columns.Count];
    //        worksheet.Cells[headerRange.Address].AutoFilter = true;

    //        package.Save();
    //    }
    //}

    public void ExportarParaExcelContabelizar(DataTable dataTable, string filePath)
    {
        if (!dataTable.Columns.Contains("Horas Decimais"))
        {
            dataTable.Columns.Add("Horas Decimais", typeof(double));
        }

        foreach (DataRow row in dataTable.Rows)
        {
            string qtdHora = row["Qtd de Hora"].ToString();
            if (TimeSpan.TryParse(qtdHora, out TimeSpan timeSpan))
            {
                double horasDecimais = timeSpan.TotalHours;
                row["Horas Decimais"] = Math.Round(horasDecimais, 2);
            }
            else
            {
                row["Horas Decimais"] = 0;
            }
        }

        var mapaPreparadores = ObterMapaPreparadores();

        if (!dataTable.Columns.Contains("N Mecanog"))
        {
            dataTable.Columns.Add("N Mecanog", typeof(string));
        }

        foreach (DataRow row in dataTable.Rows)
        {
            string preparador = row["Preparador"].ToString().Trim();
            if (mapaPreparadores.TryGetValue(preparador, out string numeroMecanografico))
            {
                row["N Mecanog"] = numeroMecanografico; 
            }
            else
            {
                row["N Mecanog"] = " "; 
            }
        }


        if (!dataTable.Columns.Contains("Mês"))
        {
            dataTable.Columns.Add("Mês", typeof(string));
        }

        foreach (DataRow row in dataTable.Rows)
        {
            if (DateTime.TryParse(row["Data da Tarefa"].ToString(), out DateTime data))
            {
                string mes = data.ToString("MMMM", new CultureInfo("pt-PT"));
                row["Mês"] = char.ToUpper(mes[0]) + mes.Substring(1);
            }
        }

        if (!dataTable.Columns.Contains("Artigo"))
        {
            dataTable.Columns.Add("Artigo", typeof(string));
        }

        if (!dataTable.Columns.Contains("Preço"))
        {
            dataTable.Columns.Add("Preço", typeof(string));
        }

        foreach (DataRow row in dataTable.Rows)
        {
            row["Artigo"] = "CM0170001";
            row["Preço"] = "23.50";
        }        

        using (var package = new ExcelPackage(new FileInfo(filePath)))
        {
            var worksheet = package.Workbook.Worksheets.FirstOrDefault(ws => ws.Name == "Registos");
            if (worksheet == null)
            {
                worksheet = package.Workbook.Worksheets.Add("Registos");
            }

            var borderColor = Color.FromArgb(155, 194, 230); // Cor da borda azul
            var thinBorderStyle = ExcelBorderStyle.Thin; // Estilo da borda fina
            var oddRowColor = Color.FromArgb(221, 235, 247); // Cor de fundo para linhas ímpares

            for (int i = 0; i < dataTable.Columns.Count; i++)
            {
                var cell = worksheet.Cells[2, i + 1];

                switch (i)
                {
                    case 0:
                        cell.Value = "Data";
                        break;
                    case 1:
                        cell.Value = "Obra";
                        break;
                    case 3:
                        cell.Value = "Codigo";
                        break;
                    case 6:
                        cell.Value = "Observações";
                        break;
                    case 7:
                        cell.Value = "Designação";
                        break;
                    default:
                        cell.Value = dataTable.Columns[i].ColumnName;
                        break;
                }

                cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                cell.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(155, 194, 230));
                cell.Style.Font.Bold = true;
                cell.Style.Font.Color.SetColor(Color.Black);
                cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                cell.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                cell.Style.Font.Size = 12;
                worksheet.Row(2).Height = 30;

                switch (i)
                {
                    case 0: // Data
                        worksheet.Column(i + 1).Width = 15;
                        break;
                    case 1: // Obra
                        worksheet.Column(i + 1).Width = 15;
                        break;
                    case 2: // Preparador
                        worksheet.Column(i + 1).Width = 20;
                        break;
                    case 3: // Codigo
                        worksheet.Column(i + 1).Width = 18;
                        break;
                    case 4: // Hora Inicial
                        worksheet.Column(i + 1).Width = 13;
                        break;
                    case 5: // Hora Final
                        worksheet.Column(i + 1).Width = 13;
                        break;
                    case 6: // Observações
                        worksheet.Column(i + 1).Width = 30;
                        break;
                    case 7: // Designação
                        worksheet.Column(i + 1).Width = 45;
                        break;
                    case 8: // Qtd de Hora
                        worksheet.Column(i + 1).Width = 14;
                        break;
                    case 9: // Horas Decimais
                        worksheet.Column(i + 1).Width = 15;
                        break;
                    case 10: // N Mecanografico
                        worksheet.Column(i + 1).Width = 15;
                        break;
                    case 11: // Mes
                        worksheet.Column(i + 1).Width = 16;
                        break;
                    case 12: // Artigo
                        worksheet.Column(i + 1).Width = 16;
                        break;
                    case 13: // Preço
                        worksheet.Column(i + 1).Width = 12;
                        break;
                   
                }
            }

            for (int row = 0; row < dataTable.Rows.Count; row++)
            {
                for (int col = 0; col < dataTable.Columns.Count; col++)
                {
                    var cell = worksheet.Cells[row + 3, col + 1];
                    string value = dataTable.Rows[row][col].ToString();

                    if (col == 6 && value == "0")
                    {
                        cell.Value = "";
                    }
                    else
                    {
                        cell.Value = value;
                    }

                    cell.Style.Border.Top.Style = thinBorderStyle;
                    cell.Style.Border.Left.Style = thinBorderStyle;
                    cell.Style.Border.Right.Style = thinBorderStyle;
                    cell.Style.Border.Bottom.Style = thinBorderStyle;

                    cell.Style.Border.Top.Color.SetColor(borderColor);
                    cell.Style.Border.Left.Color.SetColor(borderColor);
                    cell.Style.Border.Right.Color.SetColor(borderColor);
                    cell.Style.Border.Bottom.Color.SetColor(borderColor);

                    cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    cell.Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    var azulClaro = Color.FromArgb(221, 235, 247);
                    var branco = Color.White;

                    // A partir da linha 3 do Excel, alternar cor de fundo por linha
                    if ((row % 2) == 0) // linha par (0, 2, 4...) → linha 3, 5, 7... no Excel
                    {
                        cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        cell.Style.Fill.BackgroundColor.SetColor(azulClaro);
                    }
                    else // linha ímpar (1, 3, 5...) → linha 4, 6, 8... no Excel
                    {
                        cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        cell.Style.Fill.BackgroundColor.SetColor(branco);
                    }

                }
            }

            var headerRange = worksheet.Cells[2, 1, 2, dataTable.Columns.Count];
            worksheet.Cells[headerRange.Address].AutoFilter = true;

            package.Save();
        }
    }

    public Dictionary<string, string> ObterMapaPreparadores()
    {
        var mapa = new Dictionary<string, string>();
        ComunicaBD comunicaBD = new ComunicaBD();

        try
        {
            comunicaBD.ConectarBD();

            string query = "SELECT Nome, NumeroMecanografico FROM dbo.nPreparadores1";
            using (SqlCommand cmd = new SqlCommand(query, comunicaBD.GetConnection()))
            using (SqlDataReader reader = cmd.ExecuteReader())
            {
                while (reader.Read())
                {
                    string nome = reader["Nome"].ToString().Trim();
                    string numero = reader["NumeroMecanografico"].ToString().Trim();

                    if (!mapa.ContainsKey(nome))
                        mapa[nome] = numero;
                }
            }
        }
        finally
        {
            comunicaBD.DesonectarBD();
        }

        return mapa;
    }


    public void ExportarParaExcelContabelizarPorMes(DataTable dataTable, string filePath)
    {
        using (var package = new ExcelPackage(new FileInfo(filePath)))
        {
            DateTime dataDoMes = DateTime.Parse(dataTable.Rows[0]["Data da Tarefa"].ToString());
            string nomeMes = dataDoMes.ToString("MMMM yyyy");

            var worksheet = package.Workbook.Worksheets.FirstOrDefault(ws => ws.Name == nomeMes);
            if (worksheet == null)
            {
                worksheet = package.Workbook.Worksheets.Add(nomeMes);
            }

            var borderColor = Color.FromArgb(155, 194, 230); // Cor da borda azul
            var thinBorderStyle = ExcelBorderStyle.Thin; // Estilo da borda fina
            var oddRowColor = Color.FromArgb(221, 235, 247); // Cor de fundo para linhas ímpares

            for (int i = 0; i < dataTable.Columns.Count; i++)
            {
                var cell = worksheet.Cells[2, i + 1];

                switch (i)
                {
                    case 0:
                        cell.Value = "Data";
                        break;
                    case 1:
                        cell.Value = "Obra";
                        break;
                    case 3:
                        cell.Value = "Codigo";
                        break;
                    case 6:
                        cell.Value = "Observações";
                        break;
                    case 7:
                        cell.Value = "Designação";
                        break;
                    default:
                        cell.Value = dataTable.Columns[i].ColumnName;
                        break;
                }

                cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                cell.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(155, 194, 230));
                cell.Style.Font.Bold = true;
                cell.Style.Font.Color.SetColor(Color.Black);
                cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                cell.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                cell.Style.Font.Size = 12;
                worksheet.Row(2).Height = 30;

                switch (i)
                {
                    case 0: // Data
                        worksheet.Column(i + 1).Width = 15;
                        break;
                    case 1: // Obra
                        worksheet.Column(i + 1).Width = 15;
                        break;
                    case 2: // Preparador
                        worksheet.Column(i + 1).Width = 20;
                        break;
                    case 3: // Codigo
                        worksheet.Column(i + 1).Width = 18;
                        break;
                    case 4: // Hora Inicial
                        worksheet.Column(i + 1).Width = 13;
                        break;
                    case 5: // Hora Final
                        worksheet.Column(i + 1).Width = 13;
                        break;
                    case 6: // Observações
                        worksheet.Column(i + 1).Width = 25;
                        break;
                    case 7: // Designação
                        worksheet.Column(i + 1).Width = 45;
                        break;
                    case 8: // Qtd de Hora
                        worksheet.Column(i + 1).Width = 14;
                        break;
                }
            }

            for (int row = 0; row < dataTable.Rows.Count; row++)
            {
                for (int col = 0; col < dataTable.Columns.Count; col++)
                {
                    var cell = worksheet.Cells[row + 3, col + 1];
                    string value = dataTable.Rows[row][col].ToString();

                    if (col == 6 && value == "0")
                    {
                        cell.Value = "";
                    }
                    else
                    {
                        cell.Value = value;
                    }

                    cell.Style.Border.Top.Style = thinBorderStyle;
                    cell.Style.Border.Left.Style = thinBorderStyle;
                    cell.Style.Border.Right.Style = thinBorderStyle;
                    cell.Style.Border.Bottom.Style = thinBorderStyle;

                    cell.Style.Border.Top.Color.SetColor(borderColor);
                    cell.Style.Border.Left.Color.SetColor(borderColor);
                    cell.Style.Border.Right.Color.SetColor(borderColor);
                    cell.Style.Border.Bottom.Color.SetColor(borderColor);

                    cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    cell.Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    if ((row + 1) % 2 != 0)
                    {
                        cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        cell.Style.Fill.BackgroundColor.SetColor(oddRowColor);
                    }
                    if (col == 0) // Data
                    {
                        if (DateTime.TryParse(value, out DateTime dateValue))
                        {
                            cell.Value = dateValue;
                            cell.Style.Numberformat.Format = "dd/MM/yyyy"; // Formato de data
                        }
                    }
                    else if (col == 8) // Qtd de Hora
                    {
                        if (double.TryParse(value, out double hoursValue))
                        {
                            cell.Value = TimeSpan.FromHours(hoursValue);
                            cell.Style.Numberformat.Format = "[hh]:mm"; // Formato de horas
                        }
                    }
                }
            }

            var headerRange = worksheet.Cells[2, 1, 2, dataTable.Columns.Count];
            worksheet.Cells[headerRange.Address].AutoFilter = true;

            //FormatExcelColumnsByMonth(dataTable, filePath);
            package.Save();
        }
    }

    //public void FormatExcelColumns(string filePath)
    //{
    //    using (var package = new ExcelPackage(new FileInfo(filePath)))
    //    {
    //        var worksheet = package.Workbook.Worksheets[0];

    //        for (int row = 3; row <= worksheet.Dimension.End.Row; row++) 
    //        {
    //            var cell = worksheet.Cells[row, 1]; 
    //            if (DateTime.TryParse(cell.Text, out DateTime dateValue))
    //            {
    //                cell.Style.Numberformat.Format = "dd/MM/yyyy"; 
    //            }
    //        }
    //        for (int row = 3; row <= worksheet.Dimension.End.Row; row++) 
    //        {
    //            var cell = worksheet.Cells[row, 9]; 
    //            if (double.TryParse(cell.Text, out double hoursValue))
    //            {
    //                cell.Style.Numberformat.Format = "[hh]:mm"; 
    //            }
    //        }
    //        package.Save();
    //    }
    //}

    public void FormatExcelColumnsByMonth(DataTable dataTable, string filePath)
    {
        DateTime dataDoMes = DateTime.Parse(dataTable.Rows[0]["Data da Tarefa"].ToString());
        string nomeMes = dataDoMes.ToString("MMMM yyyy"); 

        using (var package = new ExcelPackage(new FileInfo(filePath)))
        {
            var worksheet = package.Workbook.Worksheets.FirstOrDefault(ws => ws.Name == nomeMes);

            if (worksheet == null)
            {
                worksheet = package.Workbook.Worksheets.Add(nomeMes);
            }

            for (int row = 3; row <= worksheet.Dimension.End.Row; row++) 
            {
                var cell = worksheet.Cells[row, 1]; 
                if (DateTime.TryParse(cell.Text, out DateTime dateValue))
                {
                    cell.Style.Numberformat.Format = "dd/MM/yyyy"; 
                }
            }

            for (int row = 3; row <= worksheet.Dimension.End.Row; row++) 
            {
                var cell = worksheet.Cells[row, 8]; 
                if (double.TryParse(cell.Text, out double hoursValue))
                {
                    cell.Style.Numberformat.Format = "[hh]:mm"; 
                }
            }            
            package.Save();
        }
    }

   
}


