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
    public partial class Add_Rubrique : Form
    {
        public int idSoc;
        ServiceDAO serv = new ServiceDAO();

        public Add_Rubrique()
        {
            InitializeComponent();
        }

        private void btn_Insertion_Click(object sender, EventArgs e)
        {
            int NbMax = serv.OrdreSeq(idSoc);
            bool rt = false;
            string CodeRub = txt_CodeRub.Text;
            string Libelles = txt_NameRub.Text.Replace("'", "''");
            NbMax = NbMax + 1;
            rt = serv.insertRubrique(CodeRub, Libelles,idSoc.ToString(),NbMax);
            if (rt)
            {
                MessageBox.Show("Adding completed successfully");
                txt_CodeRub.Text = "";
                txt_NameRub.Text = "";
                txt_CodeRub.Focus();
            }
            else
            {
                MessageBox.Show("Error of insert", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btn_Initialiser_Click(object sender, EventArgs e)
        {
            txt_CodeRub.Text = "";
            txt_NameRub.Text = "";
            txt_CodeRub.Focus();
        }
    }
}
