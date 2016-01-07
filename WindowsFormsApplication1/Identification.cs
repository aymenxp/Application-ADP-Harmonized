using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using HarmonizedApp.Manager;
using System.Data.OleDb;

namespace HarmonizedApp
{
    public partial class Identification : Form
    {
        ServiceDAO service = new ServiceDAO();

        public Identification()
        {
            InitializeComponent();
            this.txt_con.Text = "Administrateur";
            this.txt_MotDePasse.Text = "ABCSAGESQL";
        }

        private void btn_Ok_Click(object sender, EventArgs e)
        {
            string Login = string.Empty;
            string Pass = string.Empty;
            string User = string.Empty;
            int idSociete = 0;
            string CodeSoc = string.Empty;
            string NomSoc = string.Empty;
            if (this.cmb_NomSoc.SelectedIndex != -1)
            {
                idSociete = ((DataRaison)this.cmb_NomSoc.SelectedItem).Num;
                CodeSoc = ((DataRaison)this.cmb_NomSoc.SelectedItem).CodeSoc;
                NomSoc = ((DataRaison)this.cmb_NomSoc.SelectedItem).Societe;
            }
            if ((txt_con.Text == "Administrateur") && (txt_MotDePasse.Text == "ABCSAGESQL"))
            {
                Hide();
               (Application.OpenForms["Acceuil"] as Acceuil).AfficherApropos();
               (Application.OpenForms["Acceuil"] as Acceuil).Disable();
               (Application.OpenForms["Acceuil"] as Acceuil).AfficherUser(txt_con.Text,NomSoc);
               (Application.OpenForms["Acceuil"] as Acceuil).idSociete = idSociete;
               (Application.OpenForms["Acceuil"] as Acceuil).CodeSoc = CodeSoc;
               (Application.OpenForms["Acceuil"] as Acceuil).NomSoc = NomSoc;
             
            }
            else
            {
                OleDbDataReader dr = service.VerifierIdentif(txt_con.Text,txt_MotDePasse.Text);
                System.Data.DataTable dt = new System.Data.DataTable();
                dt.Load(dr);
                foreach (System.Data.DataRow row in dt.Rows)
                {
                    Login = row["Login"].ToString();
                    Pass = row["Pass"].ToString();
                }
                if ((Login == txt_con.Text) && (txt_MotDePasse.Text == Pass))
                {
                    Hide();
                    (Application.OpenForms["Acceuil"] as Acceuil).AfficherUser(Login, NomSoc);
                    (Application.OpenForms["Acceuil"] as Acceuil).Disable();
                    (Application.OpenForms["Acceuil"] as Acceuil).idSociete = idSociete;
                    (Application.OpenForms["Acceuil"] as Acceuil).CodeSoc = CodeSoc;
                    (Application.OpenForms["Acceuil"] as Acceuil).NomSoc = NomSoc;
                }
                else
                {
                    MessageBox.Show("Utilisateur Inéxistant", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    txt_con.Text = "";
                    txt_MotDePasse.Text = "";
                    txt_con.Focus();
                }
            }
        }

        public void Souris()
        {
            txt_con.Focus();
        }

        private void Identification_Load(object sender, EventArgs e)
        {
            List<DataRaison> list = service.GetSociteList();
            cmb_NomSoc.DataSource = list.ToArray();
            cmb_NomSoc.ValueMember = "societe";
            cmb_NomSoc.DisplayMember = "societe";
            if (list.Count > 0)
            {
                cmb_NomSoc.SelectedIndex = 0;
            }
            txt_con.Focus();
        }
    }
}
