using Guna.UI2.WinForms;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Web.UI.WebControls.WebParts;
using System.Windows.Forms;

namespace OfelizCM
{
    public partial class Frm_Main : Form
    {
        public string nomeUsuario2 { get; set; }

        private Point locationAddTarefasOriginal;
        private Point locationObrasOriginal;

        public Frm_Main()
        {
            InitializeComponent();  
            
        }

        private void Frm_Main_Load(object sender, EventArgs e)
        {
            DateTime horaAtual = DateTime.Now;
            string saudacao = ObterSaudacao(horaAtual);
            labelSaudacao.Text = saudacao;
            ButtaoTarefas.PerformClick();
            this.FormClosing += Frm_Main_FormClosing;
            CarregarPastasNaComboBoxAno();            
            SetUtilizadorPeloNome();        
            MostrarMensagemFesta();
            VerificarUsuario();
            VerificarEExportar();

            labelloginuser.Text = nomeUsuario2;

            locationAddTarefasOriginal = new Point(1, 230);
            locationObrasOriginal = new Point(1, 310);


        }

        public void VerificarEExportar()
        {
            
                ExportExcelRegistoparaContabelizar();
                ExportExcelRegistoparaContabelizarPorMes();            
        }                          

        private bool TemTarefasPendentes()
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

                string query = "SELECT COUNT(*) FROM dbo.RegistoTempo " +
                               "WHERE Preparador = '" + nomeFormatado + "' " +
                               "AND ([Qtd de Hora] = '00:00:00' OR [Hora Final] = '00:00:00')";

                SqlCommand command = new SqlCommand(query, comunicaBD.GetConnection());

                object result = command.ExecuteScalar();

                return Convert.ToInt32(result) > 0;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erro ao verificar tarefas pendentes: " + ex.Message);
                return false;
            }
            finally
            {
                comunicaBD.DesonectarBD();
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
                            guna2Button3.Visible = true;
                            guna2Button4.Visible = true;
                        }
                        else
                        {
                            guna2Button3.Visible = false;
                            guna2Button4.Visible = false;
                            ButtaoTarefas.Location = new Point(700, 22);
                            guna2Button2.Location = new Point(855, 22);
                            ButtonObrasabrir.Location = new Point(1030, 22);
                            guna2Button6.Location = new Point(1190, 22);
                            ComboBoxObrasPesquisaAno.Location = new Point(1350, 22);
                            ComboBoxObrasPesquisa.Location = new Point(1440, 22);
                        }
                    }
                    else
                    {
                        guna2Button3.Visible = false;
                        guna2Button4.Visible = false;
                        ButtaoTarefas.Location = new Point(700, 22);
                        guna2Button2.Location = new Point(855, 22);
                        ButtonObrasabrir.Location = new Point(1030, 22);
                        guna2Button6.Location = new Point(1190, 22);
                        ComboBoxObrasPesquisaAno.Location = new Point(1350, 22);
                        ComboBoxObrasPesquisa.Location = new Point(1440, 22);
                    }

                    nomeUsuario2 = Properties.Settings.Default.NomeUsuario;
                    labelloginuser.Text = nomeUsuario2;

                    if (nomeUsuario2 == "ofelizcmadmin" || nomeUsuario2 == "helder.silva")
                    {
                        labelloginuser.Visible = true;
                        guna2Button3.Visible = true;
                        guna2Button4.Visible = true;
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

        public void SetUtilizadorPeloNome()
        {
            string nomeUtilizador = Environment.UserName;
            List<string> list = new List<string>();
            ComunicaBD BD = new ComunicaBD();
            BD.ConectarBD();
            string query = "SELECT Nome FROM [TempoPreparacao].[dbo].[nPreparadores1] WHERE [nome.sigla] = '" + nomeUtilizador + "'";
            list = BD.Procurarbdlist(query);
            BD.DesonectarBD();

            if (list.Count > 0)
            {
                labelUtilizador.Text = list[0];
                Properties.Settings.Default.NomeUsuario = string.Empty;
                Properties.Settings.Default.Save();
            }
            else
            {
                labelUtilizador.Visible = false;

            if (Properties.Settings.Default.Login == "noLogin")
            {
                Frm_Login frmLogin = new Frm_Login();
                frmLogin.ShowDialog();
                this.Hide();
            }
            else if (Properties.Settings.Default.Login == "Login")
            {
                 
            }
          }
        }          

        private void guna2Button1_Click(object sender, EventArgs e)
        {
            Frm_TodasObras frm = new Frm_TodasObras();
            frm.Show();
        }
               
        private void guna2Button2_Click(object sender, EventArgs e)
        {
            container(new Frm_RegistoTempo());
        }

        private void container(object _form)
        {
            Form Fm = _form as Form;
            if (panel1.Controls.Count > 0)
            {
                panel1.Controls[0].Dispose();
            }
            Fm.TopLevel = false;
            Fm.FormBorderStyle = FormBorderStyle.None;
            Fm.Dock = DockStyle.Fill;
            panel1.Controls.Add(Fm);
            Fm.Show();
        }      

        static string ObterSaudacao(DateTime hora)
        {
            if (hora.Hour >= 0 && hora.Hour < 13)
            {
                return "Bom Dia";
            }
            else if (hora.Hour >= 13 && hora.Hour < 18)
            {
                return "Boa Tarde";
            }
            else if (hora.Hour == 18 && hora.Minute <= 30)
            {
                return "Boa Tarde";
            }
            else
            {
                return "Boa Noite";
            }
        }

        private void guna2Button3_Click(object sender, EventArgs e)
        {
            Frm_AdicionarTarefas frm = new Frm_AdicionarTarefas();
            frm.Show();
        }
       
        private static DateTime CalcularPascoa(int ano)
        {
            int a = ano % 19;
            int b = ano / 100;
            int c = ano % 100;
            int d = b / 4;
            int e = b % 4;
            int f = (b + 8) / 25;
            int g = (b - f + 1) / 16;
            int h = (19 * a + b - d - g + 15) % 30;
            int i = c / 4;
            int k = c % 4;
            int l = (32 + 2 * e + 2 * i - h - k) % 7;
            int m = (a + 21 - g + l) / 11;
            int mes = (h + l - 7 * m + 90) / 25;
            int dia = (h - m + 30) % 31 + 1;

            return new DateTime(ano, mes, dia);
        }

        private void MostrarMensagemFesta()
        {
            DateTime dataAtual = DateTime.Now;

            //Boas Festas
            DateTime inicioBoasFestas = new DateTime(dataAtual.Year, 11, 30);
            DateTime fimBoasFestas = new DateTime(dataAtual.Year, 12, 15);

            //Happy christmas
            DateTime inicioFelizNatal = new DateTime(dataAtual.Year, 12, 15);
            DateTime fimFelizNatal = new DateTime(dataAtual.Year, 12, 25);

            //New Year
            DateTime inicioBomAno = new DateTime(dataAtual.Year, 12, 26);
            DateTime fimBomAno = new DateTime(dataAtual.Year + 1, 1, 10);

            //25 de Abril
            DateTime inicio25Abril = new DateTime(dataAtual.Year, 4, 24);
            DateTime fimDia25Abril = new DateTime(dataAtual.Year, 4, 25);

            //1 de Maio Dia do Trabalhador
            DateTime inicioDiatrabalhador = new DateTime(dataAtual.Year, 4, 30);
            DateTime fimDiatrabalhador = new DateTime(dataAtual.Year, 5, 1);

            //Dia de Portugal
            DateTime inicioDia10Junho = new DateTime(dataAtual.Year, 6, 9);
            DateTime fimDia10Junho = new DateTime(dataAtual.Year, 6, 10);

            //15 de agosto – Dia da Assunção
            DateTime inicioDia15Agosto = new DateTime(dataAtual.Year, 8, 14);
            DateTime fimDia15Agosto = new DateTime(dataAtual.Year, 8, 15);

            //5 de outubro – Dia da República
            DateTime inicioDia5Outubro = new DateTime(dataAtual.Year, 10, 4);
            DateTime fimDia5Outubroo = new DateTime(dataAtual.Year, 10, 5);

            //Halloween
            DateTime inicioDiabruxas = new DateTime(dataAtual.Year, 10, 27);
            DateTime fimDiabruxas = new DateTime(dataAtual.Year, 10, 30);

            //1 de novembro – Dia de Todos os Santos
            DateTime inicioDia1Novembro = new DateTime(dataAtual.Year, 10, 31);
            DateTime fimDia1Novembro = new DateTime(dataAtual.Year, 11, 1);

            //1 de dezembro – Dia da Restauração da Independência
            DateTime inicioDia1Dezembro = new DateTime(dataAtual.Year, 11, 30);
            DateTime fimDia1Dezembro = new DateTime(dataAtual.Year, 12, 1);

            //8 de dezembro – Imaculada Conceição
            DateTime inicioDia8Dezembro = new DateTime(dataAtual.Year, 12, 7);
            DateTime fimDia8Dezembro = new DateTime(dataAtual.Year, 12, 8);

            //Dia de São João
            DateTime inicioSjOAO = new DateTime(dataAtual.Year, 6, 18);
            DateTime fimSjOAO = new DateTime(dataAtual.Year, 6, 25);

            //Dia de Santo António
            DateTime inicioSAntonio = new DateTime(dataAtual.Year, 6, 8);
            DateTime fimSAntonio = new DateTime(dataAtual.Year, 6, 13);

            //19 de março – Dia do Pai            
            DateTime diaPai = new DateTime(dataAtual.Year, 3, 19);

            //1 de junho – Dia da Criança
            DateTime diaCrianca = new DateTime(dataAtual.Year, 6, 1);

            // Calcular o Domingo de Páscoa 
            DateTime pascoa = CalcularPascoa(dataAtual.Year);

            // Dia de Carnaval
            DateTime inicioCarnaval = pascoa.AddDays(-47); 
            DateTime fimCarnaval = pascoa.AddDays(-53);


            if (dataAtual >= inicioBoasFestas && dataAtual <= fimBoasFestas)
            {
                MostrarMensagem("e umas Boas Festas");
            }
            else if (dataAtual >= inicioFelizNatal && dataAtual <= fimFelizNatal)
            {
                MostrarMensagem("e um Feliz Natal");
            }
            else if (dataAtual >= inicioBomAno && dataAtual <= fimBomAno)
            {
                MostrarMensagem("e um Bom Ano");
            }
            else if (dataAtual >= inicio25Abril && dataAtual <= fimDia25Abril)
            {
                MostrarMensagem("e Bom Feriado (Dia da Liberdade)");
            }
            else if (dataAtual >= inicioDiatrabalhador && dataAtual <= fimDiatrabalhador)
            {
                MostrarMensagem("e Bom Feriado (Dia do Trabalhador)");
            }
            else if (dataAtual >= inicioDia10Junho && dataAtual <= fimDia10Junho)
            {
                MostrarMensagem("e Bom Feriado (Dia de Portugal)");
            }
            else if (dataAtual >= inicioDia15Agosto && dataAtual <= fimDia15Agosto)
            {
                MostrarMensagem("e Bom Feriado (Dia da Assunção)");
            }
            else if (dataAtual >= inicioDia5Outubro && dataAtual <= fimDia5Outubroo)
            {
                MostrarMensagem("e Bom Feriado (Dia da República)");
            }
            else if (dataAtual >= inicioDia1Novembro && dataAtual <= fimDia1Novembro)
            {
                MostrarMensagem("e Bom Feriado (Dia de Todos os Santos)");
            }
            else if (dataAtual >= inicioDia1Dezembro && dataAtual <= fimDia1Dezembro)
            {
                MostrarMensagem("e Bom Feriado (Dia da Restauração da Independência)");
            }
            else if (dataAtual >= inicioDia8Dezembro && dataAtual <= fimDia8Dezembro)
            {
                MostrarMensagem("e Bom Feriado (Imaculada Conceição)");
            }
            else if (dataAtual >= inicioSjOAO && dataAtual <= fimSjOAO)
            {
                MostrarMensagem("e Bom S.João");
            }
            else if (dataAtual.Date == diaPai.Date)
            {
                MostrarMensagem("e um Feliz Dia do Pai");
            }
            else if (dataAtual.Date == diaCrianca.Date)
            {
                MostrarMensagem("e um Feliz Dia da Criança");
            }
            else if (dataAtual >= pascoa && dataAtual <= pascoa.AddDays(3))
            {
                MostrarMensagem("e uma Boa Páscoa");
            }
            else if (dataAtual >= inicioCarnaval && dataAtual <= fimCarnaval)
            {
                MostrarMensagem("e um Feliz Carnaval");
            }            
            else if (dataAtual >= inicioDiabruxas && dataAtual <= fimDiabruxas)
            {
                MostrarMensagem("e um Feliz Halloween");
            }
            else
            {
                OcultarMensagem();
            }
        }

        private void MostrarMensagem(string mensagem)
        {
            labelSaudacaoExtra.Visible = true;
            labelSaudacaoExtra.Text = mensagem;
        }

        private void OcultarMensagem()
        {
            labelSaudacaoExtra.Visible = false;
        }

        private void Frm_Main_FormClosing(object sender, FormClosingEventArgs e)
        {
            Properties.Settings.Default.Login = "noLogin";
            Properties.Settings.Default.Save();

            if (TemTarefasPendentes())
            {
                DialogResult result = MessageBox.Show(
                    "Existem tarefas em Execução. Tem certeza que deseja fechar o aplicativo?",
                    "Tarefas em Execução",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Warning
                );

                if (result == DialogResult.No)
                {
                    e.Cancel = true; 
                }
            }            
        }

        private void guna2Button4_Click(object sender, EventArgs e)
        {
            container(new Frm_Relatorio());
        }

        private void guna2Button6_Click(object sender, EventArgs e)
        {
            container(new Frm_Tarefas());
        }   

        public void CarregarPastasNaComboBoxObras()
        {
            ComunicaBD BD = new ComunicaBD();
            BD.ConectarBD();
            string anoSelecionado = ComboBoxObrasPesquisaAno.SelectedItem?.ToString();  
            if (!string.IsNullOrEmpty(anoSelecionado) && anoSelecionado.Length >= 4)
            {
                string anoPrefixo = anoSelecionado.Substring(anoSelecionado.Length - 2);

                string query = $"SELECT [Numero da Obra] FROM dbo.Orçamentação WHERE LEFT(CAST([Numero da Obra] AS VARCHAR(20)), 2) = '{anoPrefixo}'";
                List<string> list = BD.Procurarbdlist(query);
                BD.DesonectarBD();

                ComboBoxObrasPesquisa.Items.Clear();
                if (list.Count > 0)
                {
                    foreach (string nome in list)
                    {
                        ComboBoxObrasPesquisa.Items.Add(nome);
                    }
                }
                else
                {
                    MessageBox.Show("Nenhuma obra encontrada para o ano selecionado.");
                }
            }
            else
            {
                MessageBox.Show("Ano selecionado inválido.");
            }
        }

        private void container2(object _form)
        {
            Form Fm = _form as Form;
            if (panel1.Controls.Count > 0)
            {
                panel1.Controls[0].Dispose();
            }
            Fm.TopLevel = false;
            Fm.FormBorderStyle = FormBorderStyle.None;
            Fm.Dock = DockStyle.Fill;
            panel1.Controls.Add(Fm);

            if (Fm is Frm_Dashbord frmDashbord)
            {
                frmDashbord.NomeObra = ComboBoxObrasPesquisa.SelectedItem.ToString();
            }

            Fm.Show();
        }

        private void ComboBoxObrasPesquisa_SelectedIndexChanged(object sender, EventArgs e)
        {
            container2(new Frm_Dashbord());

        }

        private void guna2Button6_Click_1(object sender, EventArgs e)
        {
            ComboBoxObrasPesquisaAno.Visible = !ComboBoxObrasPesquisaAno.Visible;
            ComboBoxObrasPesquisa.Visible = !ComboBoxObrasPesquisa.Visible;
        }       
               

        private void guna2Button1_Click_1(object sender, EventArgs e)
        {
            container(new Frm_Atualizar());

        }


        public void CarregarPastasNaComboBoxAno()
        {
            string caminhoPasta = @"\\marconi\COMPANY SHARED FOLDER\OFELIZ\OFM\2.AN\2.CM\DP\1 Obras";

            if (Directory.Exists(caminhoPasta))
            {
                string[] subpastas = Directory.GetDirectories(caminhoPasta);

                ComboBoxObrasPesquisaAno.Items.Clear();

                // Regex para pastas com exatamente 4 caracteres alfanuméricos
                string pattern = @"^[A-Za-z0-9]{4}$";

                foreach (string subpasta in subpastas)
                {
                    string nomePasta = Path.GetFileName(subpasta);

                    if (Regex.IsMatch(nomePasta, pattern))
                    {
                        if (int.TryParse(nomePasta, out int ano))
                        {
                            if (ano >= 2022)
                            {
                                ComboBoxObrasPesquisaAno.Items.Add(nomePasta);
                            }
                        }
                    }
                }
            }
            else
            {
                MessageBox.Show("O caminho especificado não existe.");
            }
        }


        //public void CarregarPastasNaComboBoxAno()
        //{
        //    string caminhoPasta = @"\\marconi\COMPANY SHARED FOLDER\OFELIZ\OFM\2.AN\2.CM\DP\1 Obras";

        //    if (Directory.Exists(caminhoPasta))
        //    {
        //        string[] subpastas = Directory.GetDirectories(caminhoPasta);

        //        ComboBoxObrasPesquisaAno.Items.Clear();
        //        string pattern = @"^[A-Za-z0-9]{4}$";
        //        foreach (string subpasta in subpastas)
        //        {
        //            string nomePasta = Path.GetFileName(subpasta);

        //            if (Regex.IsMatch(nomePasta, pattern))
        //            {
        //                ComboBoxObrasPesquisaAno.Items.Add(nomePasta);                        
        //            }
        //        }
        //    }
        //    else
        //    {
        //        MessageBox.Show("O caminho especificado não existe.");
        //    }
        //}

        private void ComboBoxObrasPesquisaAno_SelectedIndexChanged(object sender, EventArgs e)
        {
            CarregarPastasNaComboBoxObras();
        }       
               
               
        //private void ExportExcelRegistoparaContabelizar()
        //{
        //    string query = @"
        //             SELECT [Data da Tarefa], [Numero da Obra], Preparador, [Codigo da Tarefa], [Hora Inicial], [Hora Final], ObservaçõesPreparador, Prioridade, [Qtd de Hora]
        //             FROM dbo.RegistoTempo
        //             WHERE [Data da Tarefa] >= DATEADD(MONTH, -7, GETDATE()) -- Seleciona apenas tarefas com até 7 meses de antiguidade
        //             ORDER BY [Data da Tarefa] ASC"; 


        //    ComunicaBD comunicaBD = new ComunicaBD();
        //    ExcelExport excelExport = new ExcelExport();

        //    SqlCommand command = new SqlCommand(query, comunicaBD.GetConnection());
        //    comunicaBD.ConectarBD();
        //    DataTable dataTable = comunicaBD.BuscarRegistros(command);
        //    string filePath = $@"C:\r\RegistrosANOTAR.xlsx";
        //    excelExport.ExportarParaExcelTodos(dataTable, filePath);
        //    comunicaBD.DesonectarBD();

        //    try
        //    {
        //        var excelApp = new Microsoft.Office.Interop.Excel.Application();
        //        var workbooks = excelApp.Workbooks.Open(filePath);
        //        var worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbooks.Sheets[1];

        //        DateTime dataLimite = DateTime.Now.AddMonths(-7); 
        //        int rowCount = worksheet.UsedRange.Rows.Count;

        //        for (int i = rowCount; i >= 1; i--) // Começar de baixo para cima ao excluir
        //        {
        //            var cellValue = worksheet.Cells[i, 1].Value; // A primeira coluna contém a data
        //            if (cellValue != null)
        //            {
        //                DateTime dataTarefa;
        //                if (DateTime.TryParse(cellValue.ToString(), out dataTarefa))
        //                {
        //                    if (dataTarefa < dataLimite)
        //                    {
        //                        worksheet.Rows[i].Delete(); // Excluir a linha se a data for maior que 7 meses
        //                    }
        //                }
        //            }
        //        }

        //        // Salvar e fechar o arquivo Excel
        //        workbooks.Save();
        //        workbooks.Close();
        //        excelApp.Quit();

        //        // Liberar recursos
        //        System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);

        //        // Tentar abrir o arquivo Excel após as alterações
        //        System.Diagnostics.Process.Start(filePath);
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show("Erro ao tentar processar o arquivo Excel: " + ex.Message, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //    }
        //}

        // Para Exportar 2 

        //private void ExportExcelRegistoparaContabelizar()
        //{
        //    string query = @"
        //                   SELECT [Data da Tarefa], [Numero da Obra], Preparador, [Codigo da Tarefa], [Hora Inicial], [Hora Final], ObservaçõesPreparador, Prioridade, [Qtd de Hora]
        //                   FROM dbo.RegistoTempo
        //                   ORDER BY [Data da Tarefa] ASC ";

        //    ComunicaBD comunicaBD = new ComunicaBD();
        //    ExcelExport excelExport = new ExcelExport();

        //    SqlCommand command = new SqlCommand(query, comunicaBD.GetConnection());
        //    comunicaBD.ConectarBD();
        //    DataTable dataTable = comunicaBD.BuscarRegistros(command);
        //    string filePath = $@"\\marconi\COMPANY SHARED FOLDER\OFELIZ\OFM\3.SP\7.DT\1.Técnico\5.CTS\Registo_Tempo-preparacao.xlsx";
        //    excelExport.ExportarParaExcelContabelizar(dataTable, filePath);
        //    comunicaBD.DesonectarBD();
        //    ExcelHelper.RemoverLinhasAntigas(filePath);

        //}            

        private void ExportExcelRegistoparaContabelizar()
        {
            string query = @"
                   SELECT [Data da Tarefa], [Numero da Obra], Preparador, [Codigo da Tarefa], [Hora Inicial], [Hora Final], ObservaçõesPreparador, Prioridade, [Qtd de Hora]
                   FROM dbo.RegistoTempo 
                   ORDER BY ID DESC";

            ComunicaBD comunicaBD = new ComunicaBD();
            ExcelExport excelExport = new ExcelExport();

            SqlCommand command = new SqlCommand(query, comunicaBD.GetConnection());
            comunicaBD.ConectarBD();
            DataTable dataTable = comunicaBD.BuscarRegistros(command);

            foreach (DataRow row in dataTable.Rows)
            {
                if (DateTime.TryParse(row["Data da Tarefa"].ToString(), out DateTime data))
                {
                    row["Data da Tarefa"] = data.ToString("dd/MM/yyyy");
                }
            }

            string filePath = $@"\\marconi\COMPANY SHARED FOLDER\OFELIZ\OFM\3.SP\7.DT\1.Técnico\5.CTS\Registo_Tempo-preparacao.xlsx";
            excelExport.ExportarParaExcelContabelizar(dataTable, filePath);


            comunicaBD.DesonectarBD();
            ExcelHelper.RemoverLinhasAntigas(filePath);

        }
               



        private void ExportExcelRegistoparaContabelizarPorMes()
        {
            string query = @"
                           SELECT [Data da Tarefa], [Numero da Obra], Preparador, [Codigo da Tarefa], [Hora Inicial], [Hora Final], ObservaçõesPreparador, Prioridade, [Qtd de Hora]
                           FROM dbo.RegistoTempo 
                           ORDER BY ID ASC";

            ComunicaBD comunicaBD = new ComunicaBD();
            ExcelExport excelExport = new ExcelExport();

            SqlCommand command = new SqlCommand(query, comunicaBD.GetConnection());
            comunicaBD.ConectarBD();
            DataTable dataTable = comunicaBD.BuscarRegistros(command);

            int currentMonth = DateTime.Now.Month;
            int currentYear = DateTime.Now.Year;

            var filteredRows = dataTable.AsEnumerable()
                                        .Where(row => DateTime.TryParse(row["Data da Tarefa"].ToString(), out DateTime data)
                                                   && data.Month == currentMonth
                                                   && data.Year == currentYear)
                                        .ToList();


            DataTable filteredDataTable = filteredRows.CopyToDataTable();

            string filePath = $@"\\marconi\COMPANY SHARED FOLDER\OFELIZ\OFM\3.SP\7.DT\1.Técnico\5.CTS\Registo de Tempo\Registo_Tempo_do_ano_{currentYear}.xlsx";

            if (!File.Exists(filePath))
            {
                string directoryPath = Path.GetDirectoryName(filePath);
                if (!Directory.Exists(directoryPath))
                {
                    Directory.CreateDirectory(directoryPath);
                    MessageBox.Show($" Foi criado o ficheiro Excel com o ano {currentYear}");
                }
            }

            excelExport.ExportarParaExcelContabelizarPorMes(filteredDataTable, filePath);

            comunicaBD.DesonectarBD();
        }

      

        private void guna2CustomGradientPanel2_Paint_1(object sender, PaintEventArgs e)
        {

        }
    }
}

