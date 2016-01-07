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
    public partial class Edit_Rubric : Form
    {

        public int Refer;
        public string CodeRub;
        public string LibellesRub;
        ServiceDAO service = new ServiceDAO();

        public Edit_Rubric()
        {
            InitializeComponent();
        }      

        private void btn_Initialiser_Click(object sender, EventArgs e)
        {
            txt_CodeRub.Text = "";
            txt_NameRub.Text = "";
            txt_CodeRub.Focus();
        }

        private void btn_Insertion_Click(object sender, EventArgs e)
        {
            string Libelles = string.Empty;
            string Code = string.Empty;
            bool rt = false;
            if ((txt_CodeRub.Text == "") || (txt_NameRub.Text == ""))
            {
                MessageBox.Show("You must complete all required fields !", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                Code = txt_CodeRub.Text;
                Libelles = txt_NameRub.Text;
                rt = service.ModifierRubric(Refer, Code, Libelles);
                if (rt)
                {
                    MessageBox.Show("Editing completed successfully");
                    Hide();
                    (Application.OpenForms["List_Of_Rubrique"] as List_Of_Rubrique).ViderGrid();
                    (Application.OpenForms["List_Of_Rubrique"] as List_Of_Rubrique).RemplirGrid();
                }
                else
                {
                    MessageBox.Show("Erreur de Modification", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void Edit_Rubric_Load(object sender, EventArgs e)
        {
            txt_CodeRub.Text = CodeRub;
            txt_NameRub.Text = LibellesRub;
        }
    }
}
