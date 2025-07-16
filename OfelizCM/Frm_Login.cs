using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OfelizCM
{
    public partial class Frm_Login : Form
    {
       public Frm_Login()
       {
            InitializeComponent();
            this.FormClosing += Frm_Login_FormClosing; 
       }

        private void Frm_Login_Load(object sender, EventArgs e)
        {
            TextBoxPassword.PasswordChar = '*';
        }

        public class LoginValidator
        {         
            public static bool ValidarLogin(string usuarioDigitado, string senhaDigitada)
            {
                bool loginValido = false;

                ComunicaBD BD = new ComunicaBD();

                try
                {
                    BD.ConectarBD();
                    string query = "SELECT COUNT(*) FROM dbo.Login WHERE [user] = @usuario AND [password] = @senha";
                    using (SqlCommand cmd = new SqlCommand(query, BD.GetConnection()))
                    {
                        cmd.Parameters.AddWithValue("@usuario", usuarioDigitado);
                        cmd.Parameters.AddWithValue("@senha", senhaDigitada);

                        int count = (int)cmd.ExecuteScalar();

                        if (count > 0)
                        {
                            loginValido = true;
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Erro ao validar login: " + ex.Message);
                }
                finally
                {
                    BD.DesonectarBD(); 
                }
                return loginValido;
            }     
        }              

        private void ButtonLogin_Click(object sender, EventArgs e)
        {
            string usuarioDigitado = TextBoxUsername.Text;
            string senhaDigitada = TextBoxPassword.Text;
            if (LoginValidator.ValidarLogin(usuarioDigitado, senhaDigitada))
            {
                Properties.Settings.Default.NomeUsuario = usuarioDigitado;
                Properties.Settings.Default.Save();

                Frm_Main frmMain = new Frm_Main();
                frmMain.nomeUsuario2 = usuarioDigitado;

                Properties.Settings.Default.Login = "Login";
                this.Hide();
            }
            else
            {
                MessageBox.Show("User ou possword incorretos. Tente novamente.", "Erro de Login", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void Frm_Login_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (Properties.Settings.Default.Login != "Login")
            {
                Application.Exit();
            }
        }






    }
}
