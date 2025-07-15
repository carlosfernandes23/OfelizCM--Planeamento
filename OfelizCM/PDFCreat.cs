using System;
using System.Drawing;
using System.Windows.Forms;
using iText.Kernel.Pdf;
using iText.Layout;
using iText.Layout.Element;
using iText.Layout.Properties;
using iText.Kernel.Colors; // Usado para trabalhar com cores no iTextSharp
using iText.Layout.Borders; // Para manipulação de bordas

namespace OfelizCM
{
    internal class PDFCreat
    {
        public class ExportDataGridViewToPdf
        {
            private DataGridView dataGridView;

            public ExportDataGridViewToPdf(DataGridView dgv)
            {
                this.dataGridView = dgv;
            }

            public void ExportToPdf(string filePath)
            {
                try
                {
                    using (PdfWriter writer = new PdfWriter(filePath))
                    {
                        using (PdfDocument pdf = new PdfDocument(writer))
                        {
                            pdf.SetDefaultPageSize(iText.Kernel.Geom.PageSize.A4.Rotate());

                            Document document = new Document(pdf);

                            string[] colunasDefinidas = new string[] { "Numero da Obra", "Nome da Obra", "Tarefa", "Preparador", "Prioridades", "Data de Inicio", "Data de Conclusão" };

                            Table table = new Table(colunasDefinidas.Length);

                            DeviceRgb headerBackgroundColor = new DeviceRgb(204, 204, 204); // Cor RGB (204, 204, 204)

                            foreach (var col in colunasDefinidas)
                            {
                                table.AddHeaderCell(new Cell().Add(new Paragraph(col))
                                    .SetFontSize(8)
                                    .SetBackgroundColor(headerBackgroundColor) // Define a cor de fundo do cabeçalho
                                    .SetBorder(Border.NO_BORDER)  // Remover borda
                                    .SetBorder(new iText.Layout.Borders.SolidBorder(0.5f)) // Definir espessura da borda (0.5f)
                                    .SetTextAlignment(TextAlignment.CENTER)); // Alinha texto ao centro
                            }

                            // Adiciona os dados das células
                            foreach (DataGridViewRow row in dataGridView.Rows)
                            {
                                if (row.IsNewRow) continue;

                                // Verifica se a coluna "Tarefa" contém "Gestão de Planeamento"
                                var tarefaCell = row.Cells["Tarefa"].FormattedValue?.ToString();
                                if (!string.IsNullOrEmpty(tarefaCell) && tarefaCell.Contains("Gestão de Planeamento"))
                                {
                                    continue; 
                                }

                                foreach (var col in colunasDefinidas)
                                {
                                    string cellValue = string.Empty;

                                    if (col == "Numero da Obra")
                                    {
                                        cellValue = row.Cells["Numero da Obra"].FormattedValue.ToString();
                                    }
                                    else if (col == "Nome da Obra")
                                    {
                                        cellValue = row.Cells["Nome da Obra"].FormattedValue.ToString();
                                    }
                                    else if (col == "Tarefa")
                                    {
                                        cellValue = row.Cells["Tarefa"].FormattedValue.ToString();
                                    }
                                    else if (col == "Preparador")
                                    {
                                        cellValue = row.Cells["Preparador"].FormattedValue.ToString();
                                    }
                                    else if (col == "Prioridades")
                                    {
                                        cellValue = row.Cells["Prioridades"].FormattedValue.ToString();
                                    }
                                    else if (col == "Data de Inicio")
                                    {
                                        cellValue = row.Cells["Data de Inicio"].FormattedValue.ToString();
                                    }
                                    else if (col == "Data de Conclusão")
                                    {
                                        cellValue = row.Cells["Data de Conclusão"].FormattedValue.ToString();
                                    }

                                    if (string.IsNullOrEmpty(cellValue))
                                    {
                                        cellValue = "N/A"; // Valor padrão caso a célula esteja vazia
                                    }

                                    // Adiciona os valores às células da tabela com bordas
                                    table.AddCell(new Cell().Add(new Paragraph(cellValue)
                                        .SetFontSize(8)) // Font menor
                                        .SetBorder(Border.NO_BORDER)  // Remover borda
                                        .SetBorder(new iText.Layout.Borders.SolidBorder(0.5f)) // Definir espessura da borda (0.5f)
                                        .SetTextAlignment(TextAlignment.CENTER)); // Alinha texto ao centro
                                }
                            }

                            // Adiciona a tabela ao documento PDF
                            document.Add(table);
                        }
                    }

                    MessageBox.Show("PDF gerado com sucesso!");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Erro ao criar o PDF: " + ex.Message);
                }
            }

            public void ExportToPdfRelaotorio(string filePath)
            {
                try
                {
                    using (PdfWriter writer = new PdfWriter(filePath))
                    {
                        using (PdfDocument pdf = new PdfDocument(writer))
                        {
                            pdf.SetDefaultPageSize(iText.Kernel.Geom.PageSize.A4.Rotate());

                            Document document = new Document(pdf);

                            string[] colunasDefinidas = new string[] { "Preparador", "Numero da Obra", "Nome da Obra", "Tarefa", "Relatorio", "Prioridades", "Data de Conclusão", "Observações do Relatorio" };

                            Table table = new Table(colunasDefinidas.Length);

                            DeviceRgb headerBackgroundColor = new DeviceRgb(204, 204, 204); // Cor RGB (204, 204, 204)

                            foreach (var col in colunasDefinidas)
                            {
                                table.AddHeaderCell(new Cell().Add(new Paragraph(col))
                                    .SetFontSize(8)
                                    .SetBackgroundColor(headerBackgroundColor) // Define a cor de fundo do cabeçalho
                                    .SetBorder(Border.NO_BORDER)  // Remover borda
                                    .SetBorder(new iText.Layout.Borders.SolidBorder(0.5f)) // Definir espessura da borda (0.5f)
                                    .SetTextAlignment(TextAlignment.CENTER)); // Alinha texto ao centro
                            }

                            // Adiciona os dados das células
                            foreach (DataGridViewRow row in dataGridView.Rows)
                            {
                                if (row.IsNewRow) continue;

                                foreach (var col in colunasDefinidas)
                                {
                                    string cellValue = string.Empty;
                                    
                                    if (col == "Preparador")
                                    {
                                        cellValue = row.Cells["Preparador"].FormattedValue.ToString();
                                    }
                                    else if (col == "Numero da Obra")
                                    {
                                        cellValue = row.Cells["Numero da Obra"].FormattedValue.ToString();
                                    }
                                    else if (col == "Nome da Obra")
                                    {
                                        cellValue = row.Cells["Nome da Obra"].FormattedValue.ToString();
                                    }
                                    else if (col == "Tarefa")
                                    {
                                        cellValue = row.Cells["Tarefa"].FormattedValue.ToString();
                                    }
                                    else if (col == "Relatorio")
                                    {
                                        cellValue = row.Cells["Relatorio"].FormattedValue.ToString();
                                    }
                                    else if (col == "Prioridades")
                                    {
                                        cellValue = row.Cells["Prioridades"].FormattedValue.ToString();
                                    }                                    
                                    else if (col == "Data de Conclusão")
                                    {
                                        cellValue = row.Cells["Data de Conclusão"].FormattedValue.ToString();
                                    }
                                    else if (col == "Observações do Relatorio")
                                    {
                                        cellValue = row.Cells["Observações do Relatorio"].FormattedValue.ToString();
                                    }

                                    if (string.IsNullOrEmpty(cellValue))
                                    {
                                        cellValue = ""; // Valor padrão caso a célula esteja vazia
                                    }

                                    // Adiciona os valores às células da tabela com bordas
                                    table.AddCell(new Cell().Add(new Paragraph(cellValue)
                                        .SetFontSize(6)) // Font menor
                                        .SetBorder(Border.NO_BORDER)  // Remover borda
                                        .SetBorder(new iText.Layout.Borders.SolidBorder(0.5f)) // Definir espessura da borda (0.5f)
                                        .SetTextAlignment(TextAlignment.CENTER)); // Alinha texto ao centro
                                }
                            }

                            // Adiciona a tabela ao documento PDF
                            document.Add(table);
                        }
                    }

                    MessageBox.Show("PDF gerado com sucesso!");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Erro ao criar o PDF: " + ex.Message);
                }
            }

            public void ExportToPdfRelaotorioConcluido(string filePath)
            {
                try
                {
                    using (PdfWriter writer = new PdfWriter(filePath))
                    {
                        using (PdfDocument pdf = new PdfDocument(writer))
                        {
                            pdf.SetDefaultPageSize(iText.Kernel.Geom.PageSize.A4.Rotate());

                            Document document = new Document(pdf);

                            string[] colunasDefinidas = new string[] { "Preparador", "Numero da Obra", "Nome da Obra", "Tarefa", "Relatorio", "Prioridades", "Conclusão Final" };

                            Table table = new Table(colunasDefinidas.Length);

                            DeviceRgb headerBackgroundColor = new DeviceRgb(204, 204, 204); // Cor RGB (204, 204, 204)

                            foreach (var col in colunasDefinidas)
                            {
                                table.AddHeaderCell(new Cell().Add(new Paragraph(col))
                                    .SetFontSize(8)
                                    .SetBackgroundColor(headerBackgroundColor) // Define a cor de fundo do cabeçalho
                                    .SetBorder(Border.NO_BORDER)  // Remover borda
                                    .SetBorder(new iText.Layout.Borders.SolidBorder(0.5f)) // Definir espessura da borda (0.5f)
                                    .SetTextAlignment(TextAlignment.CENTER)); // Alinha texto ao centro
                            }

                            // Adiciona os dados das células
                            foreach (DataGridViewRow row in dataGridView.Rows)
                            {
                                if (row.IsNewRow) continue;

                                foreach (var col in colunasDefinidas)
                                {
                                    string cellValue = string.Empty;

                                    if (col == "Preparador")
                                    {
                                        cellValue = row.Cells["Preparador"].FormattedValue.ToString();
                                    }
                                    else if (col == "Numero da Obra")
                                    {
                                        cellValue = row.Cells["Numero da Obra"].FormattedValue.ToString();
                                    }
                                    else if (col == "Nome da Obra")
                                    {
                                        cellValue = row.Cells["Nome da Obra"].FormattedValue.ToString();
                                    }
                                    else if (col == "Tarefa")
                                    {
                                        cellValue = row.Cells["Tarefa"].FormattedValue.ToString();
                                    }
                                    else if (col == "Relatorio")
                                    {
                                        cellValue = row.Cells["Relatorio"].FormattedValue.ToString();
                                    }
                                    else if (col == "Prioridades")
                                    {
                                        cellValue = row.Cells["Prioridades"].FormattedValue.ToString();
                                    }
                                    else if (col == "Conclusão Final")
                                    {
                                        cellValue = row.Cells["Data de Conclusão do user"].FormattedValue.ToString();
                                    }
                                    

                                    if (string.IsNullOrEmpty(cellValue))
                                    {
                                        cellValue = ""; // Valor padrão caso a célula esteja vazia
                                    }

                                    // Adiciona os valores às células da tabela com bordas
                                    table.AddCell(new Cell().Add(new Paragraph(cellValue)
                                        .SetFontSize(8)) // Font menor
                                        .SetBorder(Border.NO_BORDER)  // Remover borda
                                        .SetBorder(new iText.Layout.Borders.SolidBorder(0.5f)) // Definir espessura da borda (0.5f)
                                        .SetTextAlignment(TextAlignment.CENTER)); // Alinha texto ao centro
                                }
                            }

                            // Adiciona a tabela ao documento PDF
                            document.Add(table);
                        }
                    }

                    MessageBox.Show("PDF gerado com sucesso!");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Erro ao criar o PDF: " + ex.Message);
                }
            }

        }
    }
}



