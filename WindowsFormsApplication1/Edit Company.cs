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
    public partial class Edit_Company : Form
    {
        public int Refer;
        public string CompanyName;
        public string CodeCompany;
        ServiceDAO service = new ServiceDAO(); 


        public Edit_Company()
        {
            InitializeComponent();
        }

        private void Edit_Company_Load(object sender, EventArgs e)
        {
            txt_Soc.Text = CompanyName;
            txt_CodeSoc.Text = CodeCompany;
        }

        private void btn_Insertion_Click(object sender, EventArgs e)
        {
            string Nom = string.Empty;
            string Code = string.Empty;
            bool rt = false;
            if ((txt_CodeSoc.Text == "") || (txt_Soc.Text == ""))
            {
                MessageBox.Show("You must complete all required fields !", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                Nom = txt_CodeSoc.Text;
                Code = txt_Soc.Text;
                rt = service.ModifierSociete(Refer, Nom, Code);
                if (rt)
                {
                    MessageBox.Show("Editing completed successfully");
                    Hide();
                    (Application.OpenForms["Consult_Company"] as Consult_Company).ViderGrid();
                    (Application.OpenForms["Consult_Company"] as Consult_Company).RemplirGrid();
                }
                else
                {
                    MessageBox.Show("Erreur de Modification", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }
    }
}
