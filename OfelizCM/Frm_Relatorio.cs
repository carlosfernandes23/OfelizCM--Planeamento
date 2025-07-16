using iText.Kernel.Pdf.Canvas.Wmf;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static OfelizCM.Frm_Atualizar;
using static OfelizCM.PDFCreat;

namespace OfelizCM
{
    public partial class Frm_Relatorio : Form
    {
        public Frm_Relatorio()
        {
            InitializeComponent();
        }

        private void Relatorio_Load(object sender, EventArgs e)
        {
            ComunicaBDparaTabelaRelatoriodias();            
            ComunicaBDparaTabelaObeservações();
            ComunicaBDparaTabelaConluido();
            DataGridViewRelatorio.CellClick += DataGridViewRelatorio_CellClick;
            DataGridViewRelatorioOBS.CellClick += DataGridViewRelatorioOBS_CellClick;         
            DataGridViewRelatorioOBS.Scroll += DataGridViewRelatorioOBS_Scroll;
            DataGridViewRelatorio.Scroll += DataGridViewRelatorio_Scroll;            
        }

        private void DataGridViewRelatorio_Scroll(object sender, ScrollEventArgs e)
        {
            DataGridViewRelatorioOBS.FirstDisplayedScrollingRowIndex = DataGridViewRelatorio.FirstDisplayedScrollingRowIndex;
            DataGridViewRelatorioOBS.HorizontalScrollingOffset = DataGridViewRelatorio.HorizontalScrollingOffset;
        }

        private void DataGridViewRelatorioOBS_Scroll(object sender, ScrollEventArgs e)
        {
            DataGridViewRelatorio.FirstDisplayedScrollingRowIndex = DataGridViewRelatorioOBS.FirstDisplayedScrollingRowIndex;
            DataGridViewRelatorio.HorizontalScrollingOffset = DataGridViewRelatorioOBS.HorizontalScrollingOffset;
        }

        private void ComunicaBDparaTabelaRelatoriodias()
        {
            ComunicaBD comunicaBD = new ComunicaBD();
            try
            {
                comunicaBD.ConectarBD();

                DateTime hoje = DateTime.Now;
                int diasAtras = (int)hoje.DayOfWeek - (int)DayOfWeek.Sunday;
                if (diasAtras < 0) diasAtras += 7;  
                DateTime ultimoDomingo = hoje.AddDays(-diasAtras);

                string dataUltimoDomingo = ultimoDomingo.ToString("yyyy-MM-dd");
                string query = "SELECT Id, Preparador, [Numero da Obra], [Nome da Obra], Tarefa, Relatorio, Prioridades, [Data de Conclusão], Concluido, [Observações do Relatorio] " +
                               "FROM dbo.RegistoTarefas " +
                               "WHERE Relatorio IS NOT NULL " +
                               "AND Relatorio <> '0' " +
                               "AND [Data de Conclusão] >= '" + dataUltimoDomingo + "'"+
                               "ORDER BY [Numero da Obra] ASC";

                DataTable dataTable = comunicaBD.Procurarbd(query);

                DataGridViewRelatorio.DataSource = dataTable;

                foreach (DataRow row in dataTable.Rows)
                {
                    if (row["Relatorio"] != DBNull.Value)
                    {
                        string relatorio = row["Relatorio"].ToString();
                        relatorio = System.Text.RegularExpressions.Regex.Replace(relatorio, @"\s+", " ").Trim();
                        row["Relatorio"] = relatorio;
                    }
                }
                               
                DataGridViewRelatorio.ReadOnly = true;
                DataGridViewRelatorio.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
                DataGridViewRelatorio.Columns["Id"].Visible = false;
                DataGridViewRelatorio.Columns["Observações do Relatorio"].Visible = false;

                DataGridViewRelatorio.Columns["Preparador"].Width = 80;
                DataGridViewRelatorio.Columns["Numero da Obra"].Width = 80;
                DataGridViewRelatorio.Columns["Nome da Obra"].Width = 150;
                DataGridViewRelatorio.Columns["Tarefa"].Width = 450;
                DataGridViewRelatorio.Columns["Relatorio"].Width = 50;
                DataGridViewRelatorio.Columns["Prioridades"].Width = 200;
                DataGridViewRelatorio.Columns["Data de Conclusão"].Width = 90;
                DataGridViewRelatorio.Columns["Concluido"].Width = 70;
                if (DataGridViewRelatorio.Columns.Contains("Data de Conclusão"))
                {
                    DataGridViewRelatorio.Columns["Data de Conclusão"].HeaderText = "Data estipulada";
                }

                DataGridViewRelatorio.ClearSelection();

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

        private void ComunicaBDparaTabelaConluido()
        {
            ComunicaBD comunicaBD = new ComunicaBD();
            try
            {
                comunicaBD.ConectarBD();

                DateTime hoje = DateTime.Now;
                int diasAtras = (int)hoje.DayOfWeek - (int)DayOfWeek.Sunday;
                if (diasAtras < 0) diasAtras += 7;  
                DateTime ultimoDomingo = hoje.AddDays(-diasAtras);

                string dataUltimoDomingo = ultimoDomingo.ToString("yyyy-MM-dd");

                string query = "SELECT Id, Preparador, [Numero da Obra], [Nome da Obra], Tarefa, Prioridades, [Data de Conclusão do user] " +
                               "FROM dbo.RegistoTarefas " +
                               "WHERE Concluido = 1 " +
                               "AND [Data de Conclusão do user] >= '" + dataUltimoDomingo + "'" +
                               "ORDER BY [Numero da Obra] ASC";

                DataTable dataTable = comunicaBD.Procurarbd(query);               

                DataGridViewConcluido.DataSource = dataTable;
                DataGridViewConcluido.ReadOnly = true;
                DataGridViewConcluido.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                DataGridViewConcluido.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                DataGridViewConcluido.Columns["Id"].Visible = false;
                DataGridViewConcluido.ClearSelection();
                DataGridViewConcluido.AutoResizeColumns();

                if (DataGridViewConcluido.Columns.Contains("Data de Conclusão do user"))
                {
                    DataGridViewConcluido.Columns["Data de Conclusão do user"].HeaderText = "Data de Conclusão";
                }
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

        private void ComunicaBDparaTabelaObeservações()
        {
            ComunicaBD comunicaBD = new ComunicaBD();
            try
            {
                comunicaBD.ConectarBD();

                DateTime hoje = DateTime.Now;
                int diasAtras = (int)hoje.DayOfWeek - (int)DayOfWeek.Sunday;
                if (diasAtras < 0) diasAtras += 7; 
                DateTime ultimoDomingo = hoje.AddDays(-diasAtras);

                string dataUltimoDomingo = ultimoDomingo.ToString("yyyy-MM-dd");

                string query = "SELECT Id, [Data de Conclusão], [Observações do Relatorio] " +
                               "FROM dbo.RegistoTarefas " +
                               "WHERE Relatorio IS NOT NULL " +
                               "AND Relatorio <> '0' " +
                               "AND [Data de Conclusão] >= '" + dataUltimoDomingo + "'" +
                               "ORDER BY [Numero da Obra] ASC";

                DataTable dataTable = comunicaBD.Procurarbd(query);

                foreach (DataRow row in dataTable.Rows)
                {
                    if (row["Observações do Relatorio"] != DBNull.Value)
                    {
                        row["Observações do Relatorio"] = row["Observações do Relatorio"].ToString().Trim();
                    }
                }

                DataGridViewRelatorioOBS.DataSource = dataTable;
                DataGridViewRelatorioOBS.Columns["Id"].Visible = false;
                DataGridViewRelatorioOBS.Columns["Data de Conclusão"].Visible = false;
                DataGridViewRelatorioOBS.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                DataGridViewRelatorioOBS.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                DataGridViewRelatorioOBS.ClearSelection();
                DataGridViewRelatorioOBS.AutoResizeColumns();
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

        private void AtualizarObservacaoNoBD()
        {
            string novaObservacao = DataGridViewRelatorioOBS.SelectedRows[0].Cells["Observações do Relatorio"].Value?.ToString();
            string idTarefa = DataGridViewRelatorioOBS.SelectedRows[0].Cells["Id"].Value?.ToString();
            ComunicaBD comunicaBD = new ComunicaBD();
            try
            {
                comunicaBD.ConectarBD();

                string query = "UPDATE dbo.RegistoTarefas " +
                               "SET [Observações do Relatorio] = @NovaObservacao " +
                               "WHERE [ID] = @IdTarefa";

                SqlCommand cmd = new SqlCommand(query, comunicaBD.GetConnection());

                cmd.Parameters.AddWithValue("@NovaObservacao", novaObservacao);
                cmd.Parameters.AddWithValue("@IdTarefa", idTarefa);

                cmd.ExecuteNonQuery();

                MessageBox.Show("Observação atualizada com sucesso!");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erro ao atualizar observação: " + ex.Message);
            }
            finally
            {
                comunicaBD.DesonectarBD();
            }
        }

        private void DataGridViewRelatorio_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                int selectedRowIndex = e.RowIndex;
                if (selectedRowIndex < DataGridViewRelatorioOBS.Rows.Count)
                {
                    DataGridViewRelatorioOBS.ClearSelection();
                    DataGridViewRelatorioOBS.Rows[selectedRowIndex].Selected = true;
                }
            }
        }

        private void DataGridViewRelatorioOBS_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                int selectedRowIndex = e.RowIndex;
                if (selectedRowIndex < DataGridViewRelatorio.Rows.Count)
                {
                    DataGridViewRelatorio.ClearSelection();
                    DataGridViewRelatorio.Rows[selectedRowIndex].Selected = true;
                }
            }
        }

        private void ButtonConfirmarTarefa_Click(object sender, EventArgs e)
        {
            AtualizarObservacaoNoBD();
            ComunicaBDparaTabelaRelatoriodias();
            ComunicaBDparaTabelaObeservações();
        }

        private void guna2Button1_Click(object sender, EventArgs e)
        {
            AtualizarNaoComcluido();
            ComunicaBDparaTabelaConluido();
        }

        private void AtualizarNaoComcluido()
        {
            string idTarefa = DataGridViewConcluido.SelectedRows[0].Cells["Id"].Value?.ToString();
            string estado = "";
            string datadeconclusao = "2100-01-20";
            int concluido = 0;
            ComunicaBD comunicaBD = new ComunicaBD();
            try
            {
                comunicaBD.ConectarBD();

             string query = "UPDATE dbo.RegistoTarefas " +
                              "SET Estado = @Estado, [Data de Conclusão do user] = @datadeconclusao, Concluido = @concluido " +
                                "WHERE [ID] = @IdTarefa";

                SqlCommand cmd = new SqlCommand(query, comunicaBD.GetConnection());

                cmd.Parameters.AddWithValue("@IdTarefa", idTarefa);
                cmd.Parameters.AddWithValue("@Estado", estado);
                cmd.Parameters.AddWithValue("@datadeconclusao", datadeconclusao);
                cmd.Parameters.AddWithValue("@concluido", concluido);
                cmd.ExecuteNonQuery();
                MessageBox.Show("Tarefa foi Retirada!");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erro ao atualizar observação: " + ex.Message);
            }
            finally
            {
                comunicaBD.DesonectarBD();
            }
        }

        private void guna2ImageButton1_Click(object sender, EventArgs e)
        {
            string folderPath = @"C:\r";

            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            string filePath = Path.Combine(folderPath, "relatoriotarefasexpiradas.pdf");
            ExportDataGridViewToPdf export = new ExportDataGridViewToPdf(DataGridViewRelatorio);
            export.ExportToPdfRelaotorio(filePath);
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

            string filePath = Path.Combine(folderPath, "relatoriotarefasconcluidas.pdf");
            ExportDataGridViewToPdf export = new ExportDataGridViewToPdf(DataGridViewConcluido);
            export.ExportToPdfRelaotorioConcluido(filePath);
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

    }
}
