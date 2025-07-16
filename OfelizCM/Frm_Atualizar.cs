using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

namespace OfelizCM
{
    

    public partial class Frm_Atualizar : Form
    {
        public Frm_Atualizar()
        {
            InitializeComponent();
        }

        public class Prioridade
        {
            public string Prioridades { get; set; }
        }

        private void Atualizar_Load(object sender, EventArgs e)
        {
            ComunicaTabelas();
            VerificarUsuario();

            DataGridViewPreparadores.Columns["Nome"].HeaderText = "Nome do Preparador";
            DataGridViewPreparadores.Columns["NumeroMecanografico"].HeaderText = "Número Mecanográfico";
            DataGridViewPreparadores.Columns["nome.sigla"].HeaderText = "Nome do Computador";

            DataGridViewPrioridades.Columns["Prioridade"].HeaderText = "Prioridade Por Ordem";

            DataGridViewAutorizações.Columns["Nome"].HeaderText = "Nome do Preparador";
            DataGridViewAutorizações.Columns["AutorizacaoPreparador"].HeaderText = "Permissões de Admin";

            DataGridViewPassword.Columns["user"].HeaderText = "User";
            DataGridViewPassword.Columns["password"].HeaderText = "Password";

            DataGridViewPreparadores.CellClick += DataGridViewPreparadores_CellClick;
            DataGridViewPrioridades.CellClick += DataGridViewPrioridades_CellClick;
            DataGridViewTipologia.CellClick += DataGridViewTipologia_CellClick;
        }
                
        private void ComunicaTabelas()
        {
            ComunicaBDparaTabelaPreparadores();
            ComunicaBDparaTabelaPrioridades();
            ComunicaBDparaTabelaPreparadoresAutorizacao();
            ComunicaBDparaTabelaTipologia();
            ComunicaBDparaTabelaUserPassword();
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
                            ButtonConfirmarTarefa.Visible = true;
                            guna2Button1.Visible = true;
                            guna2Button4.Visible = true;
                            TextBoxNomePrep.Visible = true;
                            TextBoxMecanografico.Visible = true;
                            guna2Button3.Visible = true;
                            guna2Button2.Visible = true;
                            guna2Button5.Visible = true;
                            TextBoxPrio.Visible = true;
                            guna2ContainerControl3.Visible = true;
                            guna2HtmlLabel3.Visible = true;
                            guna2HtmlLabel4.Visible = true;
                            guna2ContainerControl4.Visible = true;

                        }
                        else
                        {
                            ButtonConfirmarTarefa.Visible = false;
                            guna2Button1.Visible = false;
                            guna2Button4.Visible = false;
                            TextBoxNomePrep.Visible = false;
                            TextBoxMecanografico.Visible = false;
                            guna2Button3.Visible = false;
                            guna2Button2.Visible = false;
                            guna2Button5.Visible = false;
                            TextBoxPrio.Visible = false;
                            guna2ContainerControl3.Visible = false;
                            guna2HtmlLabel3.Visible = false;
                            guna2HtmlLabel4.Visible = false;
                            guna2ContainerControl4.Visible = false;

                        }
                    }
                    else
                    {

                        ButtonConfirmarTarefa.Visible = false;
                        guna2Button1.Visible = false;
                        guna2Button4.Visible = false;
                        TextBoxNomePrep.Visible = false;
                        TextBoxMecanografico.Visible = false;
                        guna2Button3.Visible = false;
                        guna2Button2.Visible = false;
                        guna2Button5.Visible = false;
                        TextBoxPrio.Visible = false;
                        guna2ContainerControl3.Visible = false;
                        guna2HtmlLabel3.Visible = false;
                        guna2HtmlLabel4.Visible = false;
                        guna2ContainerControl4.Visible = false;
                    }
                    string nomeUsuario2 = Properties.Settings.Default.NomeUsuario;

                    if (nomeUsuario2 == "ofelizcmadmin" || nomeUsuario2 == "helder.silva")
                    {
                        ButtonConfirmarTarefa.Visible = true;
                        guna2Button1.Visible = true;
                        guna2Button4.Visible = true;
                        TextBoxNomePrep.Visible = true;
                        TextBoxMecanografico.Visible = true;
                        guna2Button3.Visible = true;
                        guna2Button2.Visible = true;
                        guna2Button5.Visible = true;
                        TextBoxPrio.Visible = true;
                        guna2ContainerControl3.Visible = true;
                        guna2HtmlLabel3.Visible = true;
                        guna2HtmlLabel4.Visible = true;
                        guna2ContainerControl4.Visible = true;
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

        private void ComunicaBDparaTabelaPreparadores()
        {
            ComunicaBD comunicaBD = new ComunicaBD();

            try
            {
                comunicaBD.ConectarBD();

                string query = "SELECT Nome, NumeroMecanografico, [nome.sigla] " +
                              "FROM dbo.nPreparadores1 ";
                             

                DataTable dataTable = comunicaBD.Procurarbd(query);

                DataGridViewPreparadores.DataSource = dataTable;

                DataGridViewPreparadores.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
                DataGridViewPreparadores.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
                DataGridViewPreparadores.ReadOnly = true;
                DataGridViewPreparadores.ClearSelection();
                DataGridViewPreparadores.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erro ao conectar à base de dados dos Preparadores: " + ex.Message);
            }
            finally
            {
                comunicaBD.DesonectarBD();
            }
        }             

        private void AtualizarPreparadores()
        {
            if (DataGridViewPreparadores.SelectedRows.Count > 0)
            {
                string Nome = DataGridViewPreparadores.SelectedRows[0].Cells["Nome"].Value.ToString();
                string NumeroMecanografico = TextBoxMecanografico.Text;
                string NomeSigla = GerarNomeSigla(Nome);


                string query = "UPDATE dbo.nPreparadores1 " +
                    "SET NumeroMecanografico = @NumeroMecanografico, " +
                    "[nome.sigla] = @nomesigla " +
                    "WHERE Nome = @Nome ";

                ComunicaBD BD = new ComunicaBD();

                try
                {
                    BD.ConectarBD();

                    using (SqlCommand cmd = new SqlCommand(query, BD.GetConnection()))
                    {
                        cmd.Parameters.AddWithValue("@Nome", Nome);
                        cmd.Parameters.AddWithValue("@nomesigla", NomeSigla);
                        cmd.Parameters.AddWithValue("@NumeroMecanografico", NumeroMecanografico);                     

                        cmd.ExecuteNonQuery();
                    }

                    ComunicaBDparaTabelaPreparadores();

                    MessageBox.Show($"Nome Atualizado com sucesso. \n Sigla  :{NomeSigla} \n Numero Mecanografico: {NumeroMecanografico}");
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

        private void guna2Button1_Click(object sender, EventArgs e)
        {
            AtualizarPreparadores();
        }

        public void ExcluirPreparadorSelecionada()
        {
            if (DataGridViewPreparadores.SelectedRows.Count > 0)
            {
                string Nome = DataGridViewPreparadores.SelectedRows[0].Cells["Nome"].Value.ToString();

                string query = "DELETE FROM dbo.nPreparadores1 WHERE Nome = @Nome";

                ComunicaBD BD = new ComunicaBD();

                try
                {
                    BD.ConectarBD();

                    using (SqlCommand cmd = new SqlCommand(query, BD.GetConnection()))
                    {
                        cmd.Parameters.AddWithValue("@Nome", Nome);

                        cmd.ExecuteNonQuery();
                    }

                    ComunicaBDparaTabelaPreparadores();

                    MessageBox.Show("Nome excluído com sucesso.");
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

        private void guna2Button4_Click(object sender, EventArgs e)
        {
            ExcluirPreparadorSelecionada();
        }

        public string GerarSigla(string nomeCompleto)
        {
            string[] palavras = nomeCompleto.Split(' ');

            List<string> iniciais = new List<string>();

            foreach (var palavra in palavras)
            {
                if (!string.IsNullOrEmpty(palavra))  
                {
                    iniciais.Add(palavra.Substring(0, 1).ToUpper());
                }
            }

            return string.Join("", iniciais);
        }

        public string GerarNomeSigla(string nomeCompleto)
        {
            string[] palavras = nomeCompleto.Split(' ');  
            List<string> nomeSigla = new List<string>();

            foreach (var palavra in palavras)
            {
                if (!string.IsNullOrEmpty(palavra)) 
                {
                    nomeSigla.Add(palavra.ToLower());  
                }
            }

            return string.Join(".", nomeSigla);
        }

        public void InserirNovoPreparador()
        {
            string nomePreparador = TextBoxNomePrep.Text; 
            string numeroMecanografico = TextBoxMecanografico.Text;
            string sigla = GerarSigla(nomePreparador);
            string nomeSigla = GerarNomeSigla(nomePreparador);
            string autorizacaoPreparador = "False";

            string query = "INSERT INTO dbo.nPreparadores1 (Nome, NumeroMecanografico, [nome.sigla], Sigla, AutorizacaoPreparador) VALUES (@Nome, @NumeroMecanografico, @nomesigla , @Sigla, @AutorizacaoPreparador)";

            ComunicaBD BD = new ComunicaBD();

            try
            {
                BD.ConectarBD();

                using (SqlCommand cmd = new SqlCommand(query, BD.GetConnection()))
                {
                    cmd.Parameters.AddWithValue("@Nome", nomePreparador);
                    cmd.Parameters.AddWithValue("@NumeroMecanografico", numeroMecanografico);
                    cmd.Parameters.AddWithValue("@nomesigla", nomeSigla);
                    cmd.Parameters.AddWithValue("@Sigla", sigla);
                    cmd.Parameters.AddWithValue("@AutorizacaoPreparador", autorizacaoPreparador);

                    cmd.ExecuteNonQuery();
                }

                ComunicaBDparaTabelaPreparadores();

                MessageBox.Show("Novo preparador inserido com sucesso.");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erro ao inserir novo preparador: " + ex.Message);
            }
            finally
            {
                BD.DesonectarBD();
            }
        }

        private void ButtonConfirmarTarefa_Click(object sender, EventArgs e)
        {
            InserirNovoPreparador();
        }

        private void ComunicaBDparaTabelaPrioridades()
        {
            ComunicaBD comunicaBD = new ComunicaBD();

            try
            {
                comunicaBD.ConectarBD();
                string query = "SELECT ID, Prioridade " +
                                       "FROM dbo.Prioridades";


                DataTable dataTable = comunicaBD.Procurarbd(query);

                DataGridViewPrioridades.DataSource = dataTable;

                DataGridViewPrioridades.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
                DataGridViewPrioridades.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
                DataGridViewPrioridades.Columns["ID"].Visible = false;
                DataGridViewPrioridades.ReadOnly = true;
                DataGridViewPrioridades.ClearSelection();
                DataGridViewPrioridades.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erro ao conectar à base de dados das Prioridades: " + ex.Message);
            }
            finally
            {
                comunicaBD.DesonectarBD();
            }
        }

        private void guna2Button3_Click(object sender, EventArgs e)
        {
            InserirNovoPrioridade();
        }

        public void InserirNovoPrioridade()
        {
            string Prioridade = TextBoxPrio.Text;

            string query = "INSERT INTO dbo.Prioridades (Prioridade) VALUES (@Prioridade)";

            ComunicaBD BD = new ComunicaBD();

            try
            {
                BD.ConectarBD();

                using (SqlCommand cmd = new SqlCommand(query, BD.GetConnection()))
                {
                    cmd.Parameters.AddWithValue("@Prioridade", Prioridade);
                    

                    cmd.ExecuteNonQuery();
                }

                ComunicaBDparaTabelaPrioridades();

                MessageBox.Show("Nova Prioridade inserida com sucesso.");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erro ao inserir nova prioridade: " + ex.Message);
            }
            finally
            {
                BD.DesonectarBD();
            }
        }

        private void AtualizarPrioridade()
        {
            if (DataGridViewPrioridades.SelectedRows.Count > 0)
            {
                string ID = DataGridViewPrioridades.SelectedRows[0].Cells["ID"].Value.ToString();   
                
                string Prioridade = TextBoxPrio.Text;

                string query = "UPDATE dbo.Prioridades " +
                 "SET Prioridade = @Prioridade " +  
                 "WHERE ID = @ID";

                ComunicaBD BD = new ComunicaBD();

                try
                {
                    BD.ConectarBD();

                    using (SqlCommand cmd = new SqlCommand(query, BD.GetConnection()))
                    {
                        cmd.Parameters.AddWithValue("@Prioridade", Prioridade);
                        cmd.Parameters.AddWithValue("@ID", ID);


                        cmd.ExecuteNonQuery();
                    }

                    ComunicaBDparaTabelaPrioridades();

                    MessageBox.Show($"Prioridade Atualizada com sucesso.");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Erro ao registrar a Prioridade: " + ex.Message);
                }
                finally
                {
                    BD.DesonectarBD();
                }
            }
        }

        private void guna2Button2_Click(object sender, EventArgs e)
        {
            AtualizarPrioridade();
        }

        private void guna2Button5_Click(object sender, EventArgs e)
        {
            ExcluirPrioridadeSelecionada();
        }

        public void ExcluirPrioridadeSelecionada()
        {
            if (DataGridViewPrioridades.SelectedRows.Count > 0)
            {
                string ID = DataGridViewPrioridades.SelectedRows[0].Cells["ID"].Value.ToString();

                string query = "DELETE FROM dbo.Prioridades WHERE ID = @ID";

                ComunicaBD BD = new ComunicaBD();

                try
                {
                    BD.ConectarBD();

                    using (SqlCommand cmd = new SqlCommand(query, BD.GetConnection()))
                    {
                        cmd.Parameters.AddWithValue("@ID", ID);

                        cmd.ExecuteNonQuery();
                    }

                    ComunicaBDparaTabelaPrioridades();

                    MessageBox.Show("Prioridade excluído com sucesso.");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Erro ao excluir a Prioridade: " + ex.Message);
                }
                finally
                {
                    BD.DesonectarBD();
                }
            }
        }
        
        private void DataGridViewPreparadores_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                DataGridViewRow row = DataGridViewPreparadores.Rows[e.RowIndex];

                TextBoxNomePrep.Text = row.Cells["Nome"].Value.ToString();
                TextBoxMecanografico.Text = row.Cells["NumeroMecanografico"].Value.ToString();
                
            }
        }

        private void DataGridViewPrioridades_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                DataGridViewRow row = DataGridViewPrioridades.Rows[e.RowIndex];

                TextBoxPrio.Text = row.Cells["Prioridade"].Value.ToString();

            }
        }

        private void DataGridViewTipologia_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                DataGridViewRow row = DataGridViewTipologia.Rows[e.RowIndex];

                TextBoxTipologia.Text = row.Cells["Tipologia"].Value.ToString();

            }
        }

        private void ComunicaBDparaTabelaPreparadoresAutorizacao()
        {
            ComunicaBD comunicaBD = new ComunicaBD();

            try
            {
                comunicaBD.ConectarBD();

                string query = "SELECT Nome, AutorizacaoPreparador " +
                              "FROM dbo.nPreparadores1 ";


                DataTable dataTable = comunicaBD.Procurarbd(query);

                DataGridViewAutorizações.DataSource = dataTable;
                DataGridViewAutorizações.ReadOnly = true;
                DataGridViewAutorizações.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
                DataGridViewAutorizações.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
                DataGridViewAutorizações.ReadOnly = true;
                DataGridViewAutorizações.ClearSelection();
                DataGridViewAutorizações.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erro ao conectar à base de dados AutorizacaoPreparador: " + ex.Message);
            }
            finally
            {
                comunicaBD.DesonectarBD();
            }
        }

        private void DarPermissao()
        {
            if (DataGridViewAutorizações.SelectedRows.Count > 0)
            {
                string Nome = DataGridViewAutorizações.SelectedRows[0].Cells["Nome"].Value.ToString();
                string autorizacaoPreparador = "True";

                string query = "UPDATE dbo.nPreparadores1 " +
                                     "SET AutorizacaoPreparador = @AutorizacaoPreparador " +  
                                     "WHERE Nome = @Nome";  

                ComunicaBD BD = new ComunicaBD();

                try
                {
                    BD.ConectarBD();

                    using (SqlCommand cmd = new SqlCommand(query, BD.GetConnection()))
                    {
                        cmd.Parameters.AddWithValue("@Nome", Nome);
                        cmd.Parameters.AddWithValue("@AutorizacaoPreparador", autorizacaoPreparador);

                        cmd.ExecuteNonQuery();
                    }

                    ComunicaBDparaTabelaPreparadoresAutorizacao();

                    MessageBox.Show($"Permissão dada com sucesso ao Preparador : {Nome} .");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Erro ao dar Permissões " + ex.Message);
                }
                finally
                {
                    BD.DesonectarBD();
                }
            }
        }

        private void guna2Button8_Click(object sender, EventArgs e)
        {
            DarPermissao();
            ComunicaBDparaTabelaPreparadoresAutorizacao();

        }

        private void RetirarPermissao()
        {
            if (DataGridViewAutorizações.SelectedRows.Count > 0)
            {
                string Nome = DataGridViewAutorizações.SelectedRows[0].Cells["Nome"].Value.ToString();
                string autorizacaoPreparador = "False";

                string query = "UPDATE dbo.nPreparadores1 " +
                                     "SET AutorizacaoPreparador = @AutorizacaoPreparador " +
                                     "WHERE Nome = @Nome";

                ComunicaBD BD = new ComunicaBD();

                try
                {
                    BD.ConectarBD();

                    using (SqlCommand cmd = new SqlCommand(query, BD.GetConnection()))
                    {
                        cmd.Parameters.AddWithValue("@Nome", Nome);
                        cmd.Parameters.AddWithValue("@AutorizacaoPreparador", autorizacaoPreparador);

                        cmd.ExecuteNonQuery();
                    }

                    ComunicaBDparaTabelaPreparadoresAutorizacao();

                    MessageBox.Show($"Permissão Retirada com sucesso ao Preparador : {Nome} .");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Erro ao dar Permissões " + ex.Message);
                }
                finally
                {
                    BD.DesonectarBD();
                }
            }
        }

        private void guna2Button7_Click(object sender, EventArgs e)
        {
            RetirarPermissao();
            ComunicaBDparaTabelaPreparadoresAutorizacao();

        }

        private void ComunicaBDparaTabelaTipologia()
        {
            ComunicaBD comunicaBD = new ComunicaBD();

            try
            {
                comunicaBD.ConectarBD();
                string query = "SELECT ID, Tipologia " +
                                       "FROM dbo.Tipologia";


                DataTable dataTable = comunicaBD.Procurarbd(query);

                DataGridViewTipologia.DataSource = dataTable;

                DataGridViewTipologia.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
                DataGridViewTipologia.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
                DataGridViewTipologia.Columns["ID"].Visible = false;
                DataGridViewTipologia.ReadOnly = true;
                DataGridViewTipologia.ClearSelection();
                DataGridViewTipologia.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erro ao conectar à base de dados da Tipologia: " + ex.Message);
            }
            finally
            {
                comunicaBD.DesonectarBD();
            }
        }

        public void InserirNovaTipologia()
        {
            string Tipologia = TextBoxTipologia.Text;

            string query = "INSERT INTO dbo.Tipologia (Tipologia) VALUES (@Tipologia)";

            ComunicaBD BD = new ComunicaBD();

            try
            {
                BD.ConectarBD();

                using (SqlCommand cmd = new SqlCommand(query, BD.GetConnection()))
                {
                    cmd.Parameters.AddWithValue("@Tipologia", Tipologia);

                    cmd.ExecuteNonQuery();
                }

                ComunicaBDparaTabelaTipologia();

                MessageBox.Show("Nova Tipologia inserida com sucesso.");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erro ao inserir nova Tipologia: " + ex.Message);
            }
            finally
            {
                BD.DesonectarBD();
            }
        }

        private void guna2Button10_Click(object sender, EventArgs e)
        {
            InserirNovaTipologia();
        }

        private void AtualizarTipologia()
        {
            if (DataGridViewTipologia.SelectedRows.Count > 0)
            {
                string ID = DataGridViewTipologia.SelectedRows[0].Cells["ID"].Value.ToString();
                string Tipologia = TextBoxTipologia.Text;


                string query = "UPDATE dbo.Tipologia " +
                               "SET Tipologia = @Tipologia " +
                               "WHERE ID = @ID"; 

                ComunicaBD BD = new ComunicaBD();

                try
                {
                    BD.ConectarBD();

                    using (SqlCommand cmd = new SqlCommand(query, BD.GetConnection()))
                    {
                        cmd.Parameters.AddWithValue("@ID", ID);
                        cmd.Parameters.AddWithValue("@Tipologia", Tipologia);


                        cmd.ExecuteNonQuery();
                    }

                    ComunicaBDparaTabelaTipologia();

                    MessageBox.Show($"Tipologia Atualizada com sucesso.");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Erro ao Atualizar a Tipologia: " + ex.Message);
                }
                finally
                {
                    BD.DesonectarBD();
                }
            }
        }

        private void guna2Button9_Click(object sender, EventArgs e)
        {
            AtualizarTipologia();
        }

        public void ExcluirTipologiaSelecionada()
        {
            if (DataGridViewTipologia.SelectedRows.Count > 0)
            {
                string ID = DataGridViewTipologia.SelectedRows[0].Cells["ID"].Value.ToString();

                string query = "DELETE FROM dbo.Tipologia WHERE ID = @ID";

                ComunicaBD BD = new ComunicaBD();

                try
                {
                    BD.ConectarBD();

                    using (SqlCommand cmd = new SqlCommand(query, BD.GetConnection()))
                    {
                        cmd.Parameters.AddWithValue("@ID", ID);

                        cmd.ExecuteNonQuery();
                    }

                    ComunicaBDparaTabelaTipologia();

                    MessageBox.Show("Tipologia excluído com sucesso.");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Erro ao excluir a Tipologia: " + ex.Message);
                }
                finally
                {
                    BD.DesonectarBD();
                }
            }
        }

        private void guna2Button6_Click(object sender, EventArgs e)
        {
            ExcluirTipologiaSelecionada();
        }

        private void ComunicaBDparaTabelaUserPassword()
        {
            ComunicaBD comunicaBD = new ComunicaBD();

            try
            {
                comunicaBD.ConectarBD();
                string query = "SELECT ID, [user], password " +
                                       "FROM dbo.Login";


                DataTable dataTable = comunicaBD.Procurarbd(query);

                DataGridViewPassword.DataSource = dataTable;

                DataGridViewPassword.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
                DataGridViewPassword.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
                DataGridViewPassword.Columns["ID"].Visible = false;              
                DataGridViewPassword.ReadOnly = true;
                DataGridViewPassword.ClearSelection();
                DataGridViewPassword.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erro ao conectar à base de dados da Tipologia: " + ex.Message);
            }
            finally
            {
                comunicaBD.DesonectarBD();
            }
        }
        
        public void ExcluirUserePasswordSelecionada()
        {
            if (DataGridViewPassword.SelectedRows.Count > 0)
            {
                string ID = DataGridViewPassword.SelectedRows[0].Cells["ID"].Value.ToString();

                string query = "DELETE FROM dbo.Login WHERE ID = @ID";

                ComunicaBD BD = new ComunicaBD();

                try
                {
                    BD.ConectarBD();

                    using (SqlCommand cmd = new SqlCommand(query, BD.GetConnection()))
                    {
                        cmd.Parameters.AddWithValue("@ID", ID);

                        cmd.ExecuteNonQuery();
                    }

                    ComunicaBDparaTabelaUserPassword();
                    
                    MessageBox.Show("User e Password excluído com sucesso.");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Erro ao excluir a o User e a Password: " + ex.Message);
                }
                finally
                {
                    BD.DesonectarBD();
                }
            }
        }

        private void guna2Button11_Click(object sender, EventArgs e)
        {
            ExcluirUserePasswordSelecionada();
        }

        public void InserirNovoUserePassword()
        {
            string user = TextBoxUser.Text;
            string password = TextBoxPassword.Text;

            string query = "INSERT INTO dbo.Login ([user], [password]) VALUES (@user, @password)";

            ComunicaBD BD = new ComunicaBD();

            try
            {
                BD.ConectarBD();

                using (SqlCommand cmd = new SqlCommand(query, BD.GetConnection()))
                {
                    cmd.Parameters.AddWithValue("@user", user);
                    cmd.Parameters.AddWithValue("@password", password);

                    cmd.ExecuteNonQuery();
                }

                ComunicaBDparaTabelaUserPassword();

                MessageBox.Show("Novo User e Password  inserida com sucesso.");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erro ao inserir o User e Password: " + ex.Message);
            }
            finally
            {
                BD.DesonectarBD();
            }
        }

        private void guna2Button13_Click(object sender, EventArgs e)
        {
            InserirNovoUserePassword();
        }

        private void AtualizarPassword()
        {
            if (DataGridViewPassword.SelectedRows.Count > 0)
            {
                string ID = DataGridViewPassword.SelectedRows[0].Cells["ID"].Value.ToString();
                string password = TextBoxPassword.Text;


                string query = "UPDATE dbo.Login " +
                               "SET password = @password " +
                               "WHERE ID = @ID";

                ComunicaBD BD = new ComunicaBD();

                try
                {
                    BD.ConectarBD();

                    using (SqlCommand cmd = new SqlCommand(query, BD.GetConnection()))
                    {
                        cmd.Parameters.AddWithValue("@ID", ID);
                        cmd.Parameters.AddWithValue("@password", password);


                        cmd.ExecuteNonQuery();
                    }

                    ComunicaBDparaTabelaUserPassword();

                    MessageBox.Show($"Password Atualizada com sucesso.");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Erro ao Atualizar a Password: " + ex.Message);
                }
                finally
                {
                    BD.DesonectarBD();
                }
            }
        }

        private void guna2Button12_Click(object sender, EventArgs e)
        {
            AtualizarPassword();
        }
    }
}
