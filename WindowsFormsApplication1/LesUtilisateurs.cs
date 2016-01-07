using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using HarmonizedApp.Manager;

namespace HarmonizedApp
{
    public partial class LesUtilisateurs : Form
    {

        ServiceDAO service = new ServiceDAO();

        public LesUtilisateurs()
        {
            InitializeComponent();
        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void btn_Insertion_Click(object sender, EventArgs e)
        {
            string NomUser = string.Empty;
            string Login = string.Empty;
            string Pass = string.Empty;
            string mat = string.Empty;
            string LeMat = string.Empty;
            bool rt = false;
            if ((txt_LOGIN.Text == "") || (txt_Pass.Text == ""))
            {
                MessageBox.Show("Il faut remplir tout les champs obligatoires", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                Login = txt_LOGIN.Text;
                Pass = txt_Pass.Text;
                rt = service.insertUserWKF(Login, Pass);
                if (rt)
                {
                    MessageBox.Show("Ajout effectué avec succès");
                    Vider();
                    Hide();                 
                }
                else
                {
                    MessageBox.Show("Erreur d'ajout", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

            public void Vider()
            {
                txt_LOGIN.Text="";
                txt_Pass.Text="";
                txt_LOGIN.Focus();
            }
        
            private void LesUtilisateurs_Load(object sender, EventArgs e)
            {
                txt_LOGIN.Focus();
            }

            private void btn_Initialiser_Click(object sender, EventArgs e)
            {
                Vider();
            }

        }
    }
