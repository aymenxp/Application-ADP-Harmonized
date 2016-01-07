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
    public partial class Modifier_Utilisateur : Form
    {

        ServiceDAO service = new ServiceDAO();
        public int RefUser;
        public string Login;
        public string Pass;

        public Modifier_Utilisateur()
        {
            InitializeComponent();
        }

        private void Modifier_Utilisateur_Load(object sender, EventArgs e)
        {
            txt_LOGIN.Text = Login;
            txt_Pass.Text = Pass;
        }

        public void Vider()
        {
            txt_LOGIN.Text = "";
            txt_LOGIN.Focus();
            txt_Pass.Text = "";
        }

        private void btn_Insertion_Click(object sender, EventArgs e)
        {
            string Login = string.Empty;
            string Pass = string.Empty;
            bool rt = false;
            if ((txt_LOGIN.Text == "") || (txt_Pass.Text == ""))
            {
                MessageBox.Show("Il faut remplir tout les champs obligatoires", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                Login = txt_LOGIN.Text;
                Pass = txt_Pass.Text;
                rt = service.ModifierLeUser(RefUser, Login, Pass);
                if (rt)
                {
                    MessageBox.Show("Modification effectuée avec succès");
                    Hide();
                    (Application.OpenForms["Liste_Des_Utilisateurs"] as Liste_Des_Utilisateurs).ViderGrid();
                    (Application.OpenForms["Liste_Des_Utilisateurs"] as Liste_Des_Utilisateurs).RemplirGrid();
                }
                else
                {
                    MessageBox.Show("Erreur de Modification", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }
    }
}
