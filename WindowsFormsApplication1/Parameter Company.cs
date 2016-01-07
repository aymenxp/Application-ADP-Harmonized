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
    public partial class Parameter_Company : Form
    {

        ServiceDAO serv = new ServiceDAO();

        public Parameter_Company()
        {
            InitializeComponent();
        }

        private void txt_CompanyName_TextChanged(object sender, EventArgs e)
        {

        }

        private void txt_CodeSoc_TextChanged(object sender, EventArgs e)
        {

        }

        private void btn_Insertion_Click(object sender, EventArgs e)
        {
            bool rt = false;
            string NomSoc = txt_Soc.Text.Replace("'", "''");
            string codeSoc = txt_CodeSoc.Text;
            rt = serv.insertRaison(NomSoc, codeSoc);
            if (rt)
            {
                MessageBox.Show("Adding completed successfully");
            }
            else
            {
                MessageBox.Show("Error of insert", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btn_Initialiser_Click(object sender, EventArgs e)
        {
            txt_CodeSoc.Text = "";
            txt_Soc.Text = "";
            txt_CodeSoc.Focus();
        }

        private void Parameter_Company_Load(object sender, EventArgs e)
        {
            txt_Soc.Focus();
        }
    }
}
