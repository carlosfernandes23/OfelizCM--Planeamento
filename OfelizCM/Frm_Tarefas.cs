using Guna.UI2.WinForms;
using ServiceStack.Text;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Web.UI.WebControls;
using System.Windows.Forms;
using static iText.StyledXmlParser.Jsoup.Select.Evaluator;
using static OfelizCM.Frm_Atualizar;
using static OfelizCM.PDFCreat;
using Bitmap = System.Drawing.Bitmap;

namespace OfelizCM
{

    public partial class Frm_Tarefas : Form
    {
        private DateTime horaInicio;
        private DateTime horaFim;
        private Timer timerContagemTempo;
        private Timer timerProgresso;
        private TimeSpan tempoDecorrido;
        private Timer timerContagemTempopass;

        Bitmap clock, hour, minute, second;

        private Timer VerificarTimerTarefas;
        private DateTime? tempoInvisivelInicio;

        [DllImport("user32.dll")]
        public static extern bool SetForegroundWindow(IntPtr hWnd);

        public Frm_Tarefas()
        {
            InitializeComponent();

            clock = new System.Drawing.Bitmap(".\\Imagens\\clock.png");
            hour = new System.Drawing.Bitmap(".\\Imagens\\hour.png");
            minute = new System.Drawing.Bitmap(".\\Imagens\\minute.png");
            second = new System.Drawing.Bitmap(".\\Imagens\\second.png");

            timerContagemTempo = new Timer();
            timerContagemTempo.Interval = 1000;
            timerContagemTempo.Tick += TimerContagemTempo_Tick;

            timerProgresso = new Timer();
            timerProgresso.Interval = 50;
            timerProgresso.Tick += TimerProgresso_Tick;

            ProgressBarProgressoTarefas.Minimum = 0;
            ProgressBarProgressoTarefas.Maximum = 100;
            ProgressBarProgressoTarefas.Value = 0;

            labelIndicadorPercentagem.Text = "0%";

            labelDataHoje.Text = DateTime.Now.Date.ToString("dd/MM/yyyy");
            IniciarProgresso();

            VerificarTimerTarefas = new Timer();
            VerificarTimerTarefas.Interval = 10000; 
            VerificarTimerTarefas.Tick += VerificarVisibilidade;
            VerificarTimerTarefas.Start();

        }

        private void Frm_Tarefas_Load(object sender, EventArgs e)
        {
            Horasempreacontar();
            ComunicaBDparaTabela();
            AtualizarTotalTarefas();
            ComunicaBDparaTabelaTarefasAbertasTodos();
            ConfigurarDataGridView();
            AtualizarVisibilidadeLabel();
            ConfirmarSeHelder();
            VerificarUsuario();
            SemNenhumatarefa();
            ComunicaBDparaTabelaTarefasAbertasHelder();
            timer3.Interval = 1000;
            timer3.Start();
            filtarLinha();
            label1.Text = DateTime.Now.ToString("dd/MM/yyyy");
            Timer timerVerificacao = new Timer();
            timerVerificacao.Interval = 60000; 
            timerVerificacao.Tick += TimerVerificacao_Tick;
            timerVerificacao.Start();
        }

        private void VerificarVisibilidade(object sender, EventArgs e)
        {
            if (!guna2ContainerControl1.Visible)
            {
                if (tempoInvisivelInicio == null)
                    tempoInvisivelInicio = DateTime.Now;

                if ((DateTime.Now - tempoInvisivelInicio.Value).TotalMinutes >= 5)
                {
                    VerificarTimerTarefas.Stop(); 

                    this.WindowState = FormWindowState.Normal;
                    this.Activate();
                    SetForegroundWindow(this.Handle);

                    MessageBox.Show("Nenhuma tarefa selecionada nos últimos 5 minutos!\nPor favor, inicie uma tarefa.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning);

                    tempoInvisivelInicio = DateTime.Now;
                    VerificarTimerTarefas.Start();
                }
            }
            else
            {
                tempoInvisivelInicio = null;
            }
        }

        private void Horasempreacontar()
        {
            long horaInicioTicks = Properties.Settings.Default.HoraInicioTicks;
            horaInicio = new DateTime(horaInicioTicks);
            labelHoradeInicio.Text = horaInicio.ToString("HH:mm:ss");
            long tempoDecorridoTicks = Properties.Settings.Default.TempoDecorridoTicks;

            if (tempoDecorridoTicks > 0)
            {
                tempoDecorrido = new TimeSpan(tempoDecorridoTicks);
            }
            else
            {
                tempoDecorrido = TimeSpan.Zero;
            }

            labelHoracontar1.Text = tempoDecorrido.ToString(@"hh\:mm\:ss");
            timerContagemTempo.Start();
        }

        private void timer(object sender, EventArgs e)
        {
            DateTime now = DateTime.Now;
            int Hour = now.Hour;
            int Minute = now.Minute;
            int Second = now.Second;

            Single AngleS = Second * 6;
            Single AngleM = Minute * 6 + AngleS / 60;
            Single AngleH = Hour * 30 + AngleM / 12;

            ClockBox.Image = clock;
            ClockBox.Controls.Add(HourBox);
            HourBox.Location = new Point(0, 0);
            HourBox.Image = rotateImage(hour, AngleH);
            HourBox.Controls.Add(MinBox);
            MinBox.Location = new Point(0, 0);
            MinBox.Image = rotateImage(minute, AngleM);
            MinBox.Controls.Add(SecBox);
            SecBox.Location = new Point(0, 0);
            SecBox.Image = rotateImage(second, AngleS);
        }

        private Bitmap rotateImage(Bitmap rotateme, float angle)
        {
            Bitmap rotatedImage = new Bitmap(rotateme.Width, rotateme.Height);
            using (Graphics g = Graphics.FromImage(rotatedImage))
            {

                g.TranslateTransform(rotateme.Width / 2, rotateme.Height / 2);
                g.RotateTransform(angle);
                g.TranslateTransform(-rotateme.Width / 2, -rotateme.Height / 2);
                g.DrawImage(rotateme, new Point(0, 0));

            }
            return rotatedImage;

        }

        private void TimerContagemTempo_Tick(object sender, EventArgs e)
        {
            TimeSpan tempoDecorrido = DateTime.Now - new DateTime(Properties.Settings.Default.HoraInicioTicks);
            labelHoracontar1.Text = tempoDecorrido.ToString(@"hh\:mm\:ss");
            Properties.Settings.Default.TempoDecorridoTicks = tempoDecorrido.Ticks;
            Properties.Settings.Default.Save();
        }

        private void TimerProgresso_Tick(object sender, EventArgs e)
        {
            string labelText = labelPercentagemProgressoT.Text;

            if (int.TryParse(labelText, out int percentual) && percentual >= 0 && percentual <= 100)
            {
                if (ProgressBarProgressoTarefas.Value < percentual)
                {
                    ProgressBarProgressoTarefas.Value++;

                    labelIndicadorPercentagem.Text = ProgressBarProgressoTarefas.Value.ToString() + "%";
                }
                else
                {
                    timerProgresso.Stop();
                }
            }
        }

        private void IniciarProgresso()
        {
            ProgressBarProgressoTarefas.Value = 0;
            labelIndicadorPercentagem.Text = "0%";
            timerProgresso.Start();
        }

        private void btnAtualizarProgressBar_Click(object sender, EventArgs e)
        {
            string labelText = labelPercentagemProgressoT.Text;

            if (int.TryParse(labelText, out int percentagem))
            {
                if (percentagem >= 0 && percentagem <= 100)
                {
                    ProgressBarProgressoTarefas.Value = percentagem;
                }
            }
        }

        private void ConfirmarSeHelder()
        {
            string nomePreparador = Environment.UserName;
            string nomeUsuario2 = Properties.Settings.Default.NomeUsuario;

            if (nomePreparador.Equals("helder.silva", StringComparison.OrdinalIgnoreCase) ||
                nomeUsuario2.Equals("helder.silva", StringComparison.OrdinalIgnoreCase) /*||
                //nomePreparador.Equals("carlos.alves", StringComparison.OrdinalIgnoreCase)*/)
            {
                DataGridViewTarefasAbertas2.Visible = true;
                DataGridViewTarefasAbertas2.Location = new Point(20, 240);
                DataGridViewTarefasAbertas2.Size = new Size(1960, 800);
                DataGridViewTarefasAbertas2.ColumnHeadersVisible = true;
                DataGridViewTarefasAbertas2.ReadOnly = true;
                DataGridViewTarefasAbertas2.ClearSelection();
                DataGridViewTarefas.Visible = false;
                                
            }
        }

        private void ComunicaBDparaTabela()
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
            ComunicaBD comunicaBD = new ComunicaBD();
            try
            {
                comunicaBD.ConectarBD();

                string query = "SELECT ID, [Numero da Obra], [Nome da Obra], Tarefa, Prioridades, Preparador, [Data de Inicio], [Data de Conclusão], Observações ,Concluido, Estado, [Codigo da Tarefa] " +
                               "FROM dbo.RegistoTarefas " +
                               "WHERE Concluido = 0 AND Estado = '' " +
                               "AND  Preparador = '" + nomeFormatado + "'" +
                               "ORDER BY [Data de Conclusão] ASC";

                DataTable dataTable = comunicaBD.Procurarbd(query);

                DataGridViewTarefas.DataSource = dataTable;

                DataGridViewTarefas.Columns["Concluido"].Visible = false;
                DataGridViewTarefas.Columns["Id"].Visible = false;
                DataGridViewTarefas.Columns["Estado"].Visible = false;
                DataGridViewTarefas.Columns["Codigo da Tarefa"].Visible = false;
                DataGridViewTarefas.Columns["Preparador"].Visible = false;
                DataGridViewTarefas.ReadOnly = true;
                DataGridViewTarefas.AutoResizeColumns();
                DataGridViewTarefas.Columns["Observações"].Width = 500;

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

        private void ButtonIniciarTarefa_Click(object sender, EventArgs e)
        {
            DateTime horaInicio = DateTime.Now;
            Properties.Settings.Default.HoraInicioTicks = horaInicio.Ticks;
            Properties.Settings.Default.TempoDecorridoTicks = 0;
            Properties.Settings.Default.Save();
            labelHoradeInicio.Text = horaInicio.ToString("HH:mm:ss");
            timerContagemTempo.Start();
            EnviarHoraInicioParaBaseDeDados();
            labelHoradeFim.Text = "";
            AtualizarTotalTarefas();
        }

        private void TerminarTarefa()
        {
            {
                horaFim = DateTime.Now;
                labelHoradeFim.Text = horaFim.ToString("HH:mm:ss");
                timerContagemTempo.Stop();

                Properties.Settings.Default.TempoDecorridoTicks = 0;
                Properties.Settings.Default.HoraFimTicks = horaFim.Ticks;
                Properties.Settings.Default.Save();
                SalvarHoraFimNaBD();
                AtualizarTotalTarefas();
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
                            guna2ImageButton3.Visible = true;
                        }
                        else
                        {
                            guna2ImageButton3.Visible = false;
                        }
                    }
                    else
                    {
                        guna2ImageButton3.Visible = false;
                    }

                    string nomeUsuario2 = Properties.Settings.Default.NomeUsuario;

                    if (nomeUsuario2 == "ofelizcmadmin" || nomeUsuario2 == "helder.silva")
                    {
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

        private void EnviarHoraInicioParaBaseDeDados()
        {
            string nomePreparador = GetNomePreparadorFormatado();
            string HoraInicio = GetHoraInicio(nomePreparador);
            if (HoraInicio == null) return;
            string Observações = string.IsNullOrEmpty(TextBoxObs.Text) ? "0" : TextBoxObs.Text;
            string HoraFinal = "00:00:00";
            string QtdHora = "00:00:00";
            string DatadaTarefa = DateTime.Now.ToString("dd/MM/yyyy");

            if (DataGridViewTarefas.SelectedRows.Count > 0)
            {
                var tarefaData = GetTarefaData();

                if (!InserirHoraInicioNoBanco(tarefaData, HoraInicio, HoraFinal, DatadaTarefa, QtdHora, Observações))
                {
                    MessageBox.Show("Erro ao registrar a hora de início.");
                }
                else
                {
                    //MessageBox.Show($"Hora Registada com sucesso. \n\n Hora Inicial da Tarefa : {HoraInicio}");
                }
            }
            else
            {
                MessageBox.Show("Por favor, selecione uma tarefa na tabela.");
            }
        }

        private string GetNomePreparadorFormatado()
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
            return string.Join(" ", partesComMaiusculas);
        }

        private string GetHoraInicio(string nomePreparador)
        {
            string dataAtual = DateTime.Now.ToString("dd/MM/yyyy");

            string queryHoraFinal = "SELECT TOP 1 [Hora Final], [Data da Tarefa] " +
                                    "FROM dbo.RegistoTempo " +
                                    "WHERE [Preparador] = @Preparador AND [Data da Tarefa] = @DataAtual " +
                                    "ORDER BY ID DESC";

            string HoraFinal = "08:30:00";

            ComunicaBD BD = new ComunicaBD();
            try
            {
                BD.ConectarBD();

                using (SqlCommand cmd = new SqlCommand(queryHoraFinal, BD.GetConnection()))
                {
                    cmd.Parameters.AddWithValue("@Preparador", nomePreparador);
                    cmd.Parameters.AddWithValue("@DataAtual", dataAtual);

                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            HoraFinal = reader["Hora Final"].ToString();
                        }
                        else
                        {
                            MessageBox.Show("Seu primeiro registro do dia, bom dia!");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erro ao buscar Hora Final: " + ex.Message);
            }
            finally
            {
                BD.DesonectarBD();
            }

            return HoraFinal;
        }

        private (int IdTarefa, string NumeroObra, string NomeObra, string Tarefa, string Prioridade, string Preparador, string Codigo) GetTarefaData()
        {
            var selectedRow = DataGridViewTarefas.SelectedRows[0];
            return (
                IdTarefa: Convert.ToInt32(selectedRow.Cells["Id"].Value),
                NumeroObra: selectedRow.Cells["Numero da Obra"].Value.ToString(),
                NomeObra: selectedRow.Cells["Nome da Obra"].Value.ToString(),
                Tarefa: selectedRow.Cells["Tarefa"].Value.ToString(),
                Prioridade: selectedRow.Cells["Prioridades"].Value.ToString(),
                Preparador: selectedRow.Cells["Preparador"].Value.ToString(),
                Codigo: selectedRow.Cells["Codigo da Tarefa"].Value.ToString()
            );
        }

        private bool InserirHoraInicioNoBanco((int IdTarefa, string NumeroObra, string NomeObra, string Tarefa, string Prioridade, string Preparador, string Codigo) tarefaData, string HoraInicio, string HoraFinal, string DatadaTarefa, string QtdHora, string Observações)
        {
            string query = "INSERT INTO dbo.RegistoTempo ([Numero da Obra], [Nome da Obra], Tarefa, Preparador, [Hora Inicial], [Hora Final], [Data da Tarefa], [Qtd de Hora], ObservaçõesPreparador, Prioridade, [Codigo da Tarefa]) " +
                           "VALUES (@NumeroObra, @NomeObra, @Tarefa, @TarefaPreparador, @HoraInicial, @HoraFinal, @DatadaTarefa, @QtddeHora, @ObservaçõesPreparador, @Prioridade, @CodigodaTarefa)";

            ComunicaBD BD = new ComunicaBD();
            try
            {
                BD.ConectarBD();
                using (SqlCommand cmd = new SqlCommand(query, BD.GetConnection()))
                {
                    cmd.Parameters.AddWithValue("@NumeroObra", tarefaData.NumeroObra);
                    cmd.Parameters.AddWithValue("@NomeObra", tarefaData.NomeObra);
                    cmd.Parameters.AddWithValue("@Tarefa", tarefaData.Tarefa);
                    cmd.Parameters.AddWithValue("@TarefaPreparador", tarefaData.Preparador);
                    cmd.Parameters.AddWithValue("@HoraInicial", HoraInicio);
                    cmd.Parameters.AddWithValue("@HoraFinal", HoraFinal);
                    cmd.Parameters.AddWithValue("@DatadaTarefa", DatadaTarefa);
                    cmd.Parameters.AddWithValue("@QtddeHora", QtdHora);
                    cmd.Parameters.AddWithValue("@ObservaçõesPreparador", Observações);
                    cmd.Parameters.AddWithValue("@Prioridade", tarefaData.Prioridade);
                    cmd.Parameters.AddWithValue("@CodigodaTarefa", tarefaData.Codigo);

                    cmd.ExecuteNonQuery();
                }
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erro ao registrar a hora de início: " + ex.Message);
                return false;
            }
            finally
            {
                BD.DesonectarBD();
            }
        }

        private void AtualizarEstadoTarefa()
        {
            if (DataGridViewTarefas.SelectedRows.Count > 0)
            {
                int idTarefa = Convert.ToInt32(DataGridViewTarefas.SelectedRows[0].Cells["Id"].Value);

                string Prioridades = DataGridViewTarefas.SelectedRows[0].Cells["Prioridades"].Value?.ToString();

                DialogResult result = MessageBox.Show("Tem certeza de que deseja marcar esta tarefa como concluída?", "Confirmar Tarefa", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);

                if (result == DialogResult.Yes)
                {
                    string query;

                    if (Prioridades == "4- Envio para Aprovação 2D/3D Trimble")
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
                            cmd.Parameters.AddWithValue("@IdTarefa", idTarefa);

                            cmd.ExecuteNonQuery();
                        }

                        ComunicaBDparaTabela();

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

        private void AtualizarParaPendente()
        {
            if (DataGridViewTarefas.SelectedRows.Count > 0)
            {
                int idTarefa = Convert.ToInt32(DataGridViewTarefas.SelectedRows[0].Cells["ID"].Value);

                DialogResult result = MessageBox.Show("Tem certeza de que deseja marcar esta tarefa como 'Pendente'?", "Confirmar Tarefa Pendente", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (result == DialogResult.Yes)
                {
                    string query = "UPDATE dbo.RegistoTarefas SET Estado = 'Pendente' WHERE ID = @IdTarefa";

                    ComunicaBD BD = new ComunicaBD();
                    try
                    {
                        BD.ConectarBD();

                        using (SqlCommand cmd = new SqlCommand(query, BD.GetConnection()))
                        {
                            cmd.Parameters.AddWithValue("@IdTarefa", idTarefa);

                            cmd.ExecuteNonQuery();

                            ComunicaBDparaTabela();

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
            if (DataGridViewTarefas.SelectedRows.Count > 0)
            {
                int idTarefa = Convert.ToInt32(DataGridViewTarefas.SelectedRows[0].Cells["ID"].Value);

                DialogResult result = MessageBox.Show("Tem certeza de que deseja marcar esta tarefa como 'Falta Defenições P/ OFELIZ'?", "Confirmar Tarefa Pendente", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (result == DialogResult.Yes)
                {
                    string query = "UPDATE dbo.RegistoTarefas SET Estado = 'Pendente', Observações = 'Faltam definições da parte do Feliz' WHERE ID = @IdTarefa";

                    ComunicaBD BD = new ComunicaBD();
                    try
                    {
                        BD.ConectarBD();

                        using (SqlCommand cmd = new SqlCommand(query, BD.GetConnection()))
                        {
                            cmd.Parameters.AddWithValue("@IdTarefa", idTarefa);

                            cmd.ExecuteNonQuery();

                            ComunicaBDparaTabela();

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
            if (DataGridViewTarefas.SelectedRows.Count > 0)
            {
                int idTarefa = Convert.ToInt32(DataGridViewTarefas.SelectedRows[0].Cells["ID"].Value);

                DialogResult result = MessageBox.Show("Tem certeza de que deseja marcar esta tarefa como 'Falta Defenições P/ Cliente'?", "Confirmar Tarefa Pendente", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (result == DialogResult.Yes)
                {
                    string query = "UPDATE dbo.RegistoTarefas SET Estado = 'Pendente', Observações = 'Faltam definições por parte do cliente' WHERE ID = @IdTarefa";

                    ComunicaBD BD = new ComunicaBD();
                    try
                    {
                        BD.ConectarBD();

                        using (SqlCommand cmd = new SqlCommand(query, BD.GetConnection()))
                        {
                            cmd.Parameters.AddWithValue("@IdTarefa", idTarefa);

                            cmd.ExecuteNonQuery();

                            ComunicaBDparaTabela();

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

        private void SalvarHoraFimNaBD()
        {
            if (DataGridViewTarefasAbertas.Rows.Count > 0)
            {
                var primeiraLinha = DataGridViewTarefasAbertas.Rows[0];

                if (primeiraLinha.Cells["Hora Inicial"].Value == null ||
                    primeiraLinha.Cells["Tarefa"].Value == null ||
                    primeiraLinha.Cells["Prioridade"].Value == null ||
                    primeiraLinha.Cells["Data da Tarefa"].Value == null ||
                    string.IsNullOrEmpty(labelHoradeFim.Text) ||
                    string.IsNullOrEmpty(labelDataHoje.Text))
                {
                    return;
                }
                string HoraInicio = DataGridViewTarefasAbertas.Rows[0].Cells["Hora Inicial"].Value.ToString();
                string HoraFim = labelHoradeFim.Text;
                string NomeTarefa = DataGridViewTarefasAbertas.Rows[0].Cells["Tarefa"].Value.ToString();
                string Prioridade = DataGridViewTarefasAbertas.Rows[0].Cells["Prioridade"].Value.ToString();
                string DataTarefaTerminada = labelDataHoje.Text;

                DateTime horaInicioDT = DateTime.ParseExact(HoraInicio, "HH:mm:ss", null);
                DateTime horaFimDT = DateTime.ParseExact(HoraFim, "HH:mm:ss", null);

                TimeSpan tempoDecorrido = horaFimDT - horaInicioDT;
                TimeSpan tempoAjustado = tempoDecorrido;

                bool horaAlmocoSubtraida = false;

                if (horaInicioDT.Hour < 12 || (horaInicioDT.Hour == 12 && horaInicioDT.Minute < 30))
                {
                    if (horaFimDT.Hour > 12 || (horaFimDT.Hour == 12 && horaFimDT.Minute >= 45))
                    {
                        TimeSpan subtracao = new TimeSpan(1, 30, 0);
                        tempoAjustado = tempoDecorrido - subtracao;
                        horaAlmocoSubtraida = true;
                    }
                }

                string QtdHoras = tempoAjustado.ToString(@"hh\:mm\:ss");

                DateTime dataTarefa = DateTime.ParseExact(DataGridViewTarefasAbertas.Rows[0].Cells["Data da Tarefa"].Value.ToString(), "dd/MM/yyyy", null);
                DateTime dataAtual = DateTime.Now;

                if (dataTarefa.Date != dataAtual.Date)
                {
                    HoraFim = "18:00:00";
                }

                string query = "UPDATE dbo.RegistoTempo " +
                               "SET [Hora Final] = @HoraFim, [Qtd de Hora] = @QtdHora " +
                               "WHERE [Hora Inicial] = @HoraInicial " +
                               "AND [Tarefa] = @Tarefa " +
                               "AND [Prioridade] = @Prioridade";

                ComunicaBD BD = new ComunicaBD();
                try
                {
                    BD.ConectarBD();

                    using (SqlCommand cmd = new SqlCommand(query, BD.GetConnection()))
                    {
                        cmd.Parameters.AddWithValue("@HoraFim", HoraFim);
                        cmd.Parameters.AddWithValue("@QtdHora", QtdHoras);
                        cmd.Parameters.AddWithValue("@HoraInicial", HoraInicio);
                        cmd.Parameters.AddWithValue("@Tarefa", NomeTarefa);
                        cmd.Parameters.AddWithValue("@Prioridade", Prioridade);

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

        private void ComunicaBDparaTabelaTarefasAbertasTodos()
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

                string query = "SELECT ID, [Numero da Obra], [Nome da Obra], Tarefa, Preparador, Prioridade ,[Hora Inicial], [Hora Final], [Data da Tarefa], [Qtd de Hora]" +
                               "FROM dbo.RegistoTempo " +
                               "WHERE Preparador = '" + nomeFormatado + "' " +
                               "AND [Hora Final] = '00:00:00' ";


                DataTable dataTable = BD.Procurarbd(query);

                DataGridViewTarefasAbertas.DataSource = dataTable;
                DataGridViewTarefasAbertas.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
                DataGridViewTarefasAbertas.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
                DataGridViewTarefasAbertas.Columns["ID"].Visible = false;
                DataGridViewTarefasAbertas.Columns["Preparador"].Visible = false;
                DataGridViewTarefasAbertas.Columns["Hora Final"].Visible = false;
                DataGridViewTarefasAbertas.Columns["Qtd de Hora"].Visible = false;
                DataGridViewTarefasAbertas.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erro ao conectar à base de dados: " + ex.Message);
            }
            finally
            {
                BD.DesonectarBD();
            }
        }

        private void ComunicaBDparaTabelaTarefasAbertasHelder()
        {

            ComunicaBD BD = new ComunicaBD();
            try
            {
                BD.ConectarBD();

                string query = "SELECT ID, Preparador, [Numero da Obra], [Nome da Obra], Tarefa, [Hora Inicial], [Hora Final], [Data da Tarefa], [Qtd de Hora], Prioridade " +
                               "FROM dbo.RegistoTempo " +
                               "WHERE [Hora Final] = '00:00:00' ";


                DataTable dataTable = BD.Procurarbd(query);

                DataGridViewTarefasAbertas2.DataSource = dataTable;
                DataGridViewTarefasAbertas2.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
                DataGridViewTarefasAbertas2.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
                DataGridViewTarefasAbertas2.Columns["ID"].Visible = false;
                DataGridViewTarefasAbertas2.Columns["Hora Final"].Visible = false;
                DataGridViewTarefasAbertas2.Columns["Qtd de Hora"].Visible = false;
                DataGridViewTarefasAbertas2.AutoResizeColumns();
                DataGridViewTarefasAbertas2.ReadOnly = true;
                DataGridViewTarefasAbertas2.ClearSelection();

            }
            catch (Exception ex)
            {
                MessageBox.Show("Erro ao conectar à base de dados: " + ex.Message);
            }
            finally
            {
                BD.DesonectarBD();
            }
        }                     

        private void enviarsoldadura()
        {
            if (DataGridViewTarefas.SelectedRows.Count > 0)
            {
                int idTarefa = Convert.ToInt32(DataGridViewTarefas.SelectedRows[0].Cells["Id"].Value);
                string Prioridades = DataGridViewTarefas.SelectedRows[0].Cells["Prioridades"].Value?.ToString();
                string Preparador = DataGridViewTarefas.SelectedRows[0].Cells["Preparador"].Value?.ToString();

                if (Preparador == "Helder Silva" || Preparador == "Elias Tinoco" || Prioridades == "8- Processo de soldadura" || Preparador == "Carlos Alves")
                {
                    AtualizarEstadoTarefaSoldadura("Elias Tinoco");
                    AtualizarEstadoTarefaSoldadura("Helder Silva");

                }
                else { }
            }
        }

        private void AtualizarLabelNdeTarefas()
        {
            labelTotalTarefas1.Text = (DataGridViewTarefas.Rows.Count - 1).ToString();
            labelobra1.Visible = false;
            labelobra2.Visible = false;
        }

        private void AtualizarVisibilidadeLabel()
        {
            if (DataGridViewTarefasAbertas != null && DataGridViewTarefasAbertas.Rows.Count > 1)
            {
                labelHoracontar1.Visible = true;
                labelComTarefas.Visible = true;
                guna2ContainerControl1.Visible = true;
            }
            else
            {
                labelHoracontar1.Visible = false;
                labelComTarefas.Visible = false;
                guna2ContainerControl1.Visible = false;

            }
        }

        private void SemNenhumatarefa()
        {
            if (DataGridViewTarefasAbertas.Rows.Count == 1)
            {
                labelSemTarefas.Visible = true;
            }
            else
            {
                labelSemTarefas.Visible = false;
            }
        }

        private void timer3_Tick(object sender, EventArgs e)
        {
            LabelHoras.Text = DateTime.Now.ToString("HH:mm:ss");
        }

        private void AtualizarTotalTarefas()
        {
            labelTotalTarefas1.Text = (DataGridViewTarefas.Rows.Count - 1).ToString();
            AtualizarIndicadorPercentagem();
        }

        private void AtualizarIndicadorPercentagem()
        {
            int totalTarefas = int.Parse(labelTotalTarefas1.Text);
            int percentagem = 0;

            switch (totalTarefas)
            {
                case 0:
                    percentagem = 100;
                    break;
                case 1:
                    percentagem = 98;
                    break;
                case 2:
                    percentagem = 92;
                    break;
                case 3:
                    percentagem = 88;
                    break;
                case 4:
                    percentagem = 82;
                    break;
                case 5:
                    percentagem = 75;
                    break;
                case 6:
                    percentagem = 72;
                    break;
                case 7:
                    percentagem = 68;
                    break;
                case 8:
                    percentagem = 62;
                    break;
                case 9:
                    percentagem = 57;
                    break;
                case 10:
                    percentagem = 52;
                    break;
                case 11:
                    percentagem = 47;
                    break;
                case 12:
                    percentagem = 43;
                    break;
                case 13:
                    percentagem = 38;
                    break;
                case 14:
                    percentagem = 33;
                    break;
                case 15:
                    percentagem = 29;
                    break;
                case 16:
                    percentagem = 25;
                    break;
                case 17:
                    percentagem = 21;
                    break;
                case 18:
                    percentagem = 18;
                    break;
                case 19:
                    percentagem = 17;
                    break;
                case 20:
                    percentagem = 13;
                    break;
                case 21:
                    percentagem = 10;
                    break;
                case 22:
                    percentagem = 8;
                    break;
                case 23:
                    percentagem = 5;
                    break;
                case 24:
                    percentagem = 2;
                    break;
                case 25:
                    percentagem = 0;
                    break;
                default:
                    percentagem = 0;
                    break;
            }

            labelPercentagemProgressoT.Text = percentagem.ToString();
            labelIndicadorPercentagem.Text = percentagem.ToString() + "%";
        }

        public void ExcluirTarefaSelecionada()
        {
            if (DataGridViewTarefasAbertas.Rows.Count > 0)
            {
                int idTarefa = Convert.ToInt32(DataGridViewTarefasAbertas.Rows[0].Cells["Id"].Value);

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
                            ComunicaBDparaTabela();

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

        private void guna2ImageButton3_Click_1(object sender, EventArgs e)
        {
            string numeroObra = "129000";
            int anoAtual = DateTime.Now.Year;
            string anoDoisUltimos = anoAtual.ToString().Substring(2, 2);
            string novoNumeroObra = anoDoisUltimos + numeroObra;
            TextBoxNumeroObra.Text = novoNumeroObra;
            string Tarefa = "Gestão de Planeamento";
            TextBoxTarefaExtra.Text = Tarefa;
            CarregarNomeObraPorCaminho();
        }
        
        private void EnviarTraefaExtra()
        {
            string HoraInicio = labelHoradeInicio.Text;
            string HoraFinal = "00:00:00";
            string QtdHora = "00:00:00";
            DateTime dataHoje = DateTime.Now;
            string DatadaTarefa = dataHoje.ToString("dd/MM/yyyy");
            DateTime HoraInicioDT = DateTime.ParseExact(HoraInicio, "HH:mm:ss", null);
            DateTime HoraFinalDT = DateTime.ParseExact(HoraFinal, "HH:mm:ss", null);
            TimeSpan tempoDecorrido = HoraFinalDT - HoraInicioDT;
            TimeSpan tempoAjustado = tempoDecorrido;
            string numeroobra = TextBoxNumeroObra.Text;

            if (labelNomeObra.Text.Contains("."))
            {
                labelNomeObra.Text = "OBRA ANUAL";

            }
            string nomedaobra = labelNomeObra.Text;
            string tarefa = TextBoxTarefaExtra.Text;
            string Observações = TextBoxTarefaExtra.Text;
            string Prioridade = "12- Tarefas Diversas";
            string CodigodaTarefa = "408";
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
                    cmd.Parameters.AddWithValue("@DatadaTarefa", DatadaTarefa);
                    cmd.Parameters.AddWithValue("@QtddeHora", QtdHora);
                    cmd.Parameters.AddWithValue("@ObservaçõesPreparador", Observações);
                    cmd.Parameters.AddWithValue("@Prioridade", Prioridade);
                    cmd.Parameters.AddWithValue("@CodigodaTarefa", CodigodaTarefa);

                    cmd.ExecuteNonQuery();
                }

                MessageBox.Show($" Hora Registada com sucesso. \n\n Hora Inicial da Tarefa : {HoraInicio}");

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

        public void InserirTarefaNoBD()
        {
            string numerodaObra = TextBoxNumeroObra.Text;

            if (labelNomeObra.Text.Contains("."))
            {
                labelNomeObra.Text = "OBRA ANUAL";
            }

            string tarefa = TextBoxTarefaExtra.Text;
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
            string Preparador = string.Join(" ", partesComMaiusculas);

            string Estado = " ";
            string observacoes = TextBoxTarefaExtra.Text;
            string prioridades = "12- Tarefas Diversas";

            DateTime dataHoje = DateTime.Now;
            DateTime dataInicio = dataHoje.Date;
            DateTime dataConclusao = dataHoje.Date;
            DateTime dataConclusaoUser = dataHoje.Date;
            int Concluido = 0;
            string CodigodaTarefa = "408";

            string nomeObra = labelNomeObra.Text;

            if (string.IsNullOrEmpty(numerodaObra) || string.IsNullOrEmpty(tarefa) || string.IsNullOrEmpty(Preparador) || string.IsNullOrEmpty(prioridades))
            {
                MessageBox.Show("Por favor, preencha todos os campos obrigatórios.");
                return;
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
                    cmd.Parameters.AddWithValue("@CodigodadaTarefa", CodigodaTarefa);
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
               
        private void ConcluirSoldadaduraEmailClara()
        {
            string Prioridades = DataGridViewTarefas.SelectedRows[0].Cells["Tarefa"].Value?.ToString();
            string Nomeobra = DataGridViewTarefas.SelectedRows[0].Cells["Nome da Obra"].Value?.ToString();
            string NumeroObra = DataGridViewTarefas.SelectedRows[0].Cells["Numero da Obra"].Value?.ToString();

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

        private void ConfirmarOT(string PrioridadeSomenteNumeros)
        {
            string Prioridades = DataGridViewTarefas.SelectedRows[0].Cells["Tarefa"].Value?.ToString();
            string Nomeobra = DataGridViewTarefas.SelectedRows[0].Cells["Nome da Obra"].Value?.ToString();
            string NumeroObra = DataGridViewTarefas.SelectedRows[0].Cells["Numero da Obra"].Value?.ToString();

            ComunicaBDprimavera BD = new ComunicaBDprimavera();
            try
            {
                BD.ConectarBD();

                string connectionString = "sua_string_de_conexao_aqui";

                string query = "UPDATE OFO " +
                                   "SET OFO.CDU_MTAutorizaSold = 1 " +
                                   "FROM GPR_OrdemFabricoOperacoes OFO " +
                                   "INNER JOIN GPR_ORDEMFABRICO O ON O.IDOrdemFabrico = OFO.IDOrdemFabrico " +
                                   "WHERE O.CDU_CodObra = @CodObra AND O.CDU_Fase = @Fase " +
                                   "AND O.Estado < 9 AND OFO.OperacaoProducao = '2.0005'";

                using (SqlCommand cmd = new SqlCommand(query, BD.GetConnection()))
                {
                    cmd.Parameters.AddWithValue("@CodObra", NumeroObra);
                    cmd.Parameters.AddWithValue("@Fase", PrioridadeSomenteNumeros);

                    cmd.ExecuteNonQuery();

                    int rowsAffected = cmd.ExecuteNonQuery();
                    if (rowsAffected > 0)
                    {
                        MessageBox.Show("Registos atualizados com sucesso, é possível agora criar OT de soldadura!", "Validação de dados", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        MessageBox.Show("Atenção, não foram encontrados registos, por favor verifique a obra e a fase!", "Validação de dados", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erro ao atualizar registros: " + ex.Message);
            }
        }

        private void AtualizarEstadoTarefaSoldadura(string preparador)
        {
            if (DataGridViewTarefas.SelectedRows.Count > 0)
            {
                int idTarefa = Convert.ToInt32(DataGridViewTarefas.SelectedRows[0].Cells["Id"].Value);
                string Prioridades = DataGridViewTarefas.SelectedRows[0].Cells["Prioridades"].Value?.ToString();
                string Numeroobra = DataGridViewTarefas.SelectedRows[0].Cells["Numero da Obra"].Value?.ToString();
                string Tarefa = DataGridViewTarefas.SelectedRows[0].Cells["Tarefa"].Value?.ToString();
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

        private void ConfigurarDataGridView()
        {
            DataGridViewTarefasAbertas.Enabled = false;
            DataGridViewTarefasAbertas.ColumnHeadersVisible = false;
            DataGridViewTarefasAbertas.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            DataGridViewTarefasAbertas.RowHeadersVisible = false;
            DataGridViewTarefasAbertas.ReadOnly = true;
        }

        private void ComunicaBDparaTabelaTarefaAbertaparaObra()
        {
            string numeroObra = string.Empty;

            if (string.IsNullOrWhiteSpace(TextBoxNumeroObra.Text))
            {
                if (DataGridViewTarefasAbertas.Rows.Count > 0 &&
                    DataGridViewTarefasAbertas.Rows[0].Cells["Numero da Obra"].Value != null)

                {
                    numeroObra = DataGridViewTarefasAbertas.Rows[0].Cells["Numero da Obra"].Value.ToString();
                }
                else
                {
                    MessageBox.Show("Por favor insira o Número de Obra.");
                    return;
                }
            }
            else
            {
                numeroObra = TextBoxNumeroObra.Text;
            }

            labelobra1.Visible = true;
            labelobra1.Text = numeroObra;
            labelobra2.Visible = true;
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

            ComunicaBD comunicaBD = new ComunicaBD();
            try
            {
                comunicaBD.ConectarBD();

                string query = "SELECT ID, [Numero da Obra], [Nome da Obra], Tarefa, Prioridades, Preparador, [Data de Inicio], [Data de Conclusão], Observações ,Concluido, Estado, [Codigo da Tarefa] " +
                               "FROM dbo.RegistoTarefas " +
                               "WHERE Concluido = 0 AND Estado = '' " +
                               "AND  Preparador = '" + nomeFormatado + "'" +
                               "AND  [Numero da Obra] = '" + numeroObra + "'" +
                               "ORDER BY [Data de Conclusão] ASC";

                DataTable dataTable = comunicaBD.Procurarbd(query);

                DataGridViewTarefas.DataSource = dataTable;

                DataGridViewTarefas.Columns["Concluido"].Visible = false;
                DataGridViewTarefas.Columns["Id"].Visible = false;
                DataGridViewTarefas.Columns["Estado"].Visible = false;
                DataGridViewTarefas.Columns["Codigo da Tarefa"].Visible = false;
                DataGridViewTarefas.Columns["Preparador"].Visible = false;
                DataGridViewTarefas.ReadOnly = true;
                DataGridViewTarefas.AutoResizeColumns();
                DataGridViewTarefas.Columns["Observações"].Width = 500;

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

        private int lastSelectedRowIndex = -1;

        private void BloquearLinha()
        {
            if (DataGridViewTarefas.SelectedRows.Count > 0)
            {
                lastSelectedRowIndex = DataGridViewTarefas.SelectedRows[0].Index;
            }

            DataGridViewTarefas.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            DataGridViewTarefas.Enabled = false;
        }

        private void DesbloquearLinha()
        {
            DataGridViewTarefas.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            DataGridViewTarefas.Enabled = true;

            if (lastSelectedRowIndex >= 0 && lastSelectedRowIndex < DataGridViewTarefas.Rows.Count)
            {
                DataGridViewTarefas.Rows[lastSelectedRowIndex].Selected = true;
            }
        }

        private void DataGridViewTarefasAbertas_SelectionChanged(object sender, EventArgs e)
        {
            if (DataGridViewTarefasAbertas.SelectedRows.Count > 0)
            {
                Properties.Settings.Default.Ultimalinhaselecionada = DataGridViewTarefasAbertas.SelectedRows[0].Index;
                Properties.Settings.Default.Save();
            }
        }

        private void guna2ImageButton16_Click(object sender, EventArgs e)
        {
            LockDataGridView();
        }

        private void LockDataGridView()
        {
            BloquearLinha();
            Properties.Settings.Default.tabelabolqueada = true;
            Properties.Settings.Default.Save();
            DataGridViewTarefas.Refresh(); 
        }

        private void UnlockDataGridView()
        {
            DesbloquearLinha();
            Properties.Settings.Default.tabelabolqueada = false;
            Properties.Settings.Default.Save();
            DataGridViewTarefas.Refresh(); 
        }

        private void filtarLinha()
        {
            if (DataGridViewTarefasAbertas.Rows.Count == 2)
            {
                SelecionarLinhaPorValores();
                guna2ImageButton16.PerformClick();
            }
            else
            { }
        }        
       
        private void guna2Button1_Click(object sender, EventArgs e)
        {
            if (DataGridViewTarefas.SelectedRows.Count > 0)
            {
                ConcluirSoldadaduraEmailClara();
                enviarsoldadura();
                AtualizarEstadoTarefa();

                if (DataGridViewTarefasAbertas.Rows.Count >= 0)
                {
                    TerminarTarefa();
                }
                else
                { }
                ComunicaBDparaTabelaTarefasAbertasTodos();
                AtualizarVisibilidadeLabel();
                AtualizarLabelNdeTarefas();
                SemNenhumatarefa();
                UnlockDataGridView();
                DataGridViewTarefas.ClearSelection();
            }
            else
            {
                MessageBox.Show("Selecione uma Tarefa para Concluir ");
            }
        }

        private void guna2Button2_Click(object sender, EventArgs e)
        {
            if (DataGridViewTarefas.SelectedRows.Count > 0)
            {
                string Prioridade = DataGridViewTarefas.SelectedRows[0].Cells["Prioridades"].Value?.ToString();

                if (Prioridade == "4- Envio para Aprovação 2D/3D Trimble" || Prioridade == "11- Desenhos de Montagem")
                {
                    ExcluirTarefaSelecionadaDoconcluido();
                }
                else
                {
                    AtualizarParaPendente();
                    TerminarTarefa();
                    ComunicaBDparaTabelaTarefasAbertasTodos();
                    AtualizarVisibilidadeLabel();
                    AtualizarLabelNdeTarefas();
                    SemNenhumatarefa();
                    UnlockDataGridView();
                    DataGridViewTarefas.ClearSelection();
                }
            }
            else
            {
                MessageBox.Show("Selecione uma Tarefa");
            }
        }

        private void guna2Button3_Click(object sender, EventArgs e)
        {
            if (DataGridViewTarefas.SelectedRows.Count > 0)
            {
                string Prioridade = DataGridViewTarefas.SelectedRows[0].Cells["Prioridades"].Value?.ToString();

                if (Prioridade == "4- Envio para Aprovação 2D/3D Trimble" || Prioridade == "11- Desenhos de Montagem")
                {
                    ExcluirTarefaSelecionadaDoconcluido();
                }
                else
                {
                    AtualizarParaPendenteDefeniçoesOfeliz();
                    TerminarTarefa();
                    ComunicaBDparaTabelaTarefasAbertasTodos();
                    AtualizarVisibilidadeLabel();
                    AtualizarLabelNdeTarefas();
                    SemNenhumatarefa();
                    UnlockDataGridView();
                    DataGridViewTarefas.ClearSelection();
                }
            }
            else
            {
                MessageBox.Show("Selecione uma Tarefa");

            }
        }

        private void guna2Button4_Click(object sender, EventArgs e)
        {
            if (DataGridViewTarefas.SelectedRows.Count > 0)
            {
                string Prioridade = DataGridViewTarefas.SelectedRows[0].Cells["Prioridades"].Value?.ToString();

                if (Prioridade == "4- Envio para Aprovação 2D/3D Trimble" || Prioridade == "11- Desenhos de Montagem")
                {
                    ExcluirTarefaSelecionadaDoconcluido();
                }
                else
                {
                    AtualizarParaPendenteDefeniçoesCliente();
                    TerminarTarefa();
                    ComunicaBDparaTabelaTarefasAbertasTodos();
                    AtualizarVisibilidadeLabel();
                    AtualizarLabelNdeTarefas();
                    SemNenhumatarefa();
                    UnlockDataGridView();
                    DataGridViewTarefas.ClearSelection();
                }
            }
            else
            {
                MessageBox.Show("Selecione uma Tarefa");

            }
        }

        private void guna2Button5_Click(object sender, EventArgs e)
        {
            ComunicaBDparaTabela();
            AtualizarLabelNdeTarefas();
            filtarLinha();
        }

        private void guna2Button6_Click(object sender, EventArgs e)
        {
            string folderPath = @"C:\r";

            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }

            string filePath = Path.Combine(folderPath, "relatorioaguardaraprovacao.pdf");

            ExportDataGridViewToPdf export = new ExportDataGridViewToPdf(DataGridViewTarefas);
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

        private void guna2Button7_Click(object sender, EventArgs e)
        {
            ComunicaBDparaTabelaTarefaAbertaparaObra();
            labelTotalTarefas1.Text = (DataGridViewTarefas.Rows.Count - 1).ToString();
            filtarLinha();
        }

        private void guna2Button8_Click(object sender, EventArgs e)
        {
            int numLinhas = DataGridViewTarefasAbertas.Rows.Count;

            if (numLinhas > 1)
            {
                MessageBox.Show("Já existe um registro de tempo aberto. Feche um registro antes de abrir outro.", "Registro de Tempo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            else
            {
                DateTime horaInicio = DateTime.Now;
                Properties.Settings.Default.HoraInicioTicks = horaInicio.Ticks;
                Properties.Settings.Default.TempoDecorridoTicks = 0;
                Properties.Settings.Default.Save();
                labelHoradeInicio.Text = horaInicio.ToString("HH:mm:ss");
                timerContagemTempo.Start();
                EnviarHoraInicioParaBaseDeDados();
                labelHoradeFim.Text = "";
                AtualizarTotalTarefas();
                ComunicaBDparaTabelaTarefasAbertasTodos();
                AtualizarVisibilidadeLabel();
                SemNenhumatarefa();
                AtualizarLabelNdeTarefas();
                filtarLinha();
                LockDataGridView();

            }
        }

        private void guna2Button9_Click(object sender, EventArgs e)
        {
            TerminarTarefa();
            ComunicaBDparaTabelaTarefasAbertasTodos();
            AtualizarVisibilidadeLabel();
            AtualizarLabelNdeTarefas();
            SemNenhumatarefa();
            UnlockDataGridView();
            DataGridViewTarefas.ClearSelection();
        }

        private void guna2Button10_Click(object sender, EventArgs e)
        {
            ExcluirTarefaSelecionada();
            ComunicaBDparaTabelaTarefasAbertasTodos();
            AtualizarVisibilidadeLabel();
            AtualizarLabelNdeTarefas();
            SemNenhumatarefa();
            UnlockDataGridView();
            DataGridViewTarefas.ClearSelection();
        }

        private void guna2Button11_Click(object sender, EventArgs e)
        {
            int numLinhas = DataGridViewTarefasAbertas.Rows.Count;

            if (numLinhas > 1)
            {
                MessageBox.Show("Já existe um registro de tempo aberto. Feche um registro antes de abrir outro.", "Registro de Tempo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            else
            {
                DateTime horaInicio = DateTime.Now;
                Properties.Settings.Default.HoraInicioTicks = horaInicio.Ticks;
                Properties.Settings.Default.TempoDecorridoTicks = 0;
                Properties.Settings.Default.Save();
                labelHoradeInicio.Text = horaInicio.ToString("HH:mm:ss");
                timerContagemTempo.Start();
                CarregarNomeObraPorCaminho();
                InserirTarefaNoBD();
                guna2Button5.PerformClick();
                EnviarTraefaExtra();
                labelHoradeFim.Text = "";
                AtualizarTotalTarefas();
                ComunicaBDparaTabelaTarefasAbertasTodos();
                AtualizarVisibilidadeLabel();
                SemNenhumatarefa();
                AtualizarLabelNdeTarefas();
                filtarLinha();
                LockDataGridView();
            }
        }

        private void SelecionarLinhaPorValores()
        {
            string numeroobra1 = DataGridViewTarefasAbertas.Rows[0].Cells["Numero da Obra"].Value.ToString().Trim();
            string tarefa = DataGridViewTarefasAbertas.Rows[0].Cells["Tarefa"].Value.ToString().Trim();
            string prioridade = DataGridViewTarefasAbertas.Rows[0].Cells["Prioridade"].Value.ToString().Trim();
            DataGridViewTarefas.ClearSelection();

            foreach (DataGridViewRow row in DataGridViewTarefas.Rows)
            {
                if (row.IsNewRow) continue;

                string numeroobra = row.Cells["Numero da Obra"].Value?.ToString();
                string tarefaNaLinha = row.Cells["Tarefa"].Value?.ToString();
                string prioridades = row.Cells["Prioridades"].Value?.ToString();

               if (numeroobra == numeroobra1 && tarefaNaLinha == tarefa && prioridades == prioridade)
               {
                      row.Selected = true;
                      DataGridViewTarefas.CurrentCell = row.Cells[1];
                      DataGridViewTarefas.FirstDisplayedScrollingRowIndex = row.Index;
                      break;
               }
            }
        }

        private void DataGridViewTarefas_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.RowIndex == lastSelectedRowIndex)
            {
                if (Properties.Settings.Default.tabelabolqueada)
                {
                    e.CellStyle.BackColor = Color.Green;
                    e.CellStyle.SelectionBackColor = Color.Green; 
                }
                else
                {
                    e.CellStyle.BackColor = DataGridViewTarefas.DefaultCellStyle.BackColor;
                }
            }
            else
            {
                e.CellStyle.BackColor = DataGridViewTarefas.DefaultCellStyle.BackColor;
            }
        }

        public void ExcluirTarefaSelecionadaDoconcluido()
        {
            int idTarefa = 0;

            if (DataGridViewTarefas.SelectedRows.Count > 0)
            {
                idTarefa = Convert.ToInt32(DataGridViewTarefas.SelectedRows[0].Cells["Id"].Value);
            }            

            if (idTarefa != 0)
            {
                DialogResult result = MessageBox.Show("Tem certeza de que deseja colocar a tarefa em pendente ?", "Confirmar", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);

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


                        MessageBox.Show("Tarefa em Pendente com sucesso.");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Erro ao colocar a tarefa em pendente : " + ex.Message);
                    }
                    finally
                    {
                        BD.DesonectarBD();
                    }
                }
            }
            else
            {
                MessageBox.Show("Por favor, selecione uma tarefa para pendente.");
            }
        }

        bool tarefaAtualizadaHoje = false;

        private void TimerVerificacao_Tick(object sender, EventArgs e)
        {
            DateTime agora = DateTime.Now;

            if (agora.DayOfWeek == DayOfWeek.Friday && agora.Hour == 17 && agora.Minute == 45 && !tarefaAtualizadaHoje)
            {
                AtualizarTarefasAtrasadas();
                tarefaAtualizadaHoje = true; 
            }

            if (agora.DayOfWeek != DayOfWeek.Friday)
            {
                tarefaAtualizadaHoje = false;
            }
        }


        private void AtualizarTarefasAtrasadas()
        {
            DateTime dataHoje = DateTime.Now.Date;
            DateTime novaData = dataHoje.AddDays(4);

            string query = @"
                            UPDATE dbo.RegistoTarefas
                            SET [Data de Conclusão] = @NovaData
                            WHERE [Data de Conclusão] <= @DataHoje
                            AND Prioridades = '8- Processo de soldadura'
                            AND Concluido = 0";

            ComunicaBD BD = new ComunicaBD();
            try
            {
                BD.ConectarBD();

                using (SqlCommand cmd = new SqlCommand(query, BD.GetConnection()))
                {
                    cmd.Parameters.AddWithValue("@NovaData", novaData);
                    cmd.Parameters.AddWithValue("@DataHoje", dataHoje);

                    int linhasAfetadas = cmd.ExecuteNonQuery();

                    MessageBox.Show($"{linhasAfetadas} tarefa(s) atualizada(s) com nova data de conclusão.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erro ao atualizar tarefas: " + ex.Message);
            }
            finally
            {
                BD.DesonectarBD();
            }
        }
      

    }
}


