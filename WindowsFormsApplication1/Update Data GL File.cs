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
    public partial class Update_Data_GL_File : Form
    {
        public Update_Data_GL_File()
        {
            InitializeComponent();
        }
        ServiceDAO service = new ServiceDAO();
        public int Refer;
        public string CodeRub;
        public string LibRub;
        public string CodeCpt;
        public string Libelles;
        public string Sens;

        private void Update_Data_GL_File_Load(object sender, EventArgs e)
        {
            txt_CodeRub.Text = CodeRub;
            textBox1.Text = LibRub;
            txt_CodeCpt.Text = CodeCpt;
            txt_Desc.Text = Libelles;
            if (Sens == "C")
            {
                cmb_sens.Text = "Crédit";
            }
            else
            {
                cmb_sens.Text = "Débit";
            }

        }

        public void Initialiser()
        {
            txt_CodeRub.Text = "";
            textBox1.Text = "";
            txt_CodeCpt.Text = "";
            txt_Desc.Text = "";
            cmb_sens.Text = "";
            txt_CodeRub.Focus();
        }

        private void btn_Initialiser_Click(object sender, EventArgs e)
        {
            Initialiser();
        }

        private void btn_Insertion_Click(object sender, EventArgs e)
        {
            string Libelles = string.Empty;
            string Code = string.Empty;
            string CodeCpt = string.Empty;
            string Sens = string.Empty;
            string LibCpt = string.Empty;
            bool rt = false;
            if ((txt_CodeRub.Text == "") || (txt_CodeCpt.Text == "")||(txt_Desc.Text==""))
            {
                MessageBox.Show("You must complete all required fields !", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                Code = txt_CodeRub.Text;
                LibCpt = textBox1.Text;
                CodeCpt = txt_CodeCpt.Text;
                Libelles = txt_Desc.Text;
                Sens = cmb_sens.Text.Substring(0, 1);
                rt = service.ModifierGlFile(Refer, Code,LibCpt, CodeCpt,Libelles,Sens);
                if (rt)
                {
                    MessageBox.Show("Editing completed successfully");
                    Hide();
                    (Application.OpenForms["List_GL_File"] as List_GL_File).ViderGrid();
                    (Application.OpenForms["List_GL_File"] as List_GL_File).RemplirGrid();
                }
                else
                {
                    MessageBox.Show("Erreur de Modification", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

    }
}
