using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using HarmonizedApp.Manager;
using System.IO;

namespace HarmonizedApp
{
    public partial class Account_Plan_about_GL_File : Form
    {

        ServiceDAO service = new ServiceDAO();
        public int IdSoc = 0;

        public Account_Plan_about_GL_File()
        {
            InitializeComponent();
        }

        private void btn_Initialiser_Click(object sender, EventArgs e)
        {
            Initialiser();
        }

        public void Initialiser()
        {
            txt_CodeRub.Text = "";
            txt_CodeCpt.Text = "";
            txt_libCode.Text = "";
            txt_Desc.Text = "";
            cmb_sens.Text = "";
            txt_CodeRub.Focus();
        }

        private void btn_Insertion_Click(object sender, EventArgs e)
        {
            string CodeRub = string.Empty;
            string CodeCpt = string.Empty;
            string Desc = string.Empty;
            string DescCodeR = string.Empty;
            string Sens = string.Empty;
            bool rt = false;
            if (!checkBox1.Checked)
            {
                if ((txt_CodeRub.Text == "") && (txt_CodeCpt.Text == "") && (txt_Desc.Text == "") && (cmb_sens.Text == ""))
                {
                    MessageBox.Show("Missing Data, please check your error", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    CodeRub = txt_CodeRub.Text;
                    DescCodeR = txt_libCode.Text;
                    CodeCpt = txt_CodeCpt.Text;
                    Desc = txt_Desc.Text;
                    Sens = cmb_sens.Text.Substring(0, 1);
                    rt = service.insertParamGlFile(CodeRub,DescCodeR, CodeCpt, Desc, Sens, IdSoc);
                }
                if (rt)
                {
                    MessageBox.Show("Insert realized with Success");
                    Initialiser();
                }
                else
                {
                    MessageBox.Show("Insert Error", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                for (int i = 0; i < dataGridView1.RowCount; i++)
                {
                    string code = dataGridView1.Rows[i].Cells[0].Value.ToString();
                    string LibCodRub = dataGridView1.Rows[i].Cells[1].Value.ToString();
                    CodeRub = dataGridView1.Rows[i].Cells[2].Value.ToString();
                    string lib = dataGridView1.Rows[i].Cells[3].Value.ToString();                 
                    Sens = dataGridView1.Rows[i].Cells[4].Value.ToString();
                    rt = service.insertParamGlFile(code,LibCodRub, CodeRub, lib, Sens, IdSoc);
                }
                if (rt)
                {
                    MessageBox.Show("Insert realized with Success");
                    Initialiser();
                }
                else
                {
                    MessageBox.Show("Insert Error", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void btn_Emplacement_Click(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();
            string file;
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Title = "Choisir le fichier à importer";
            // Extension par défaut			
            openFileDialog1.DefaultExt = "csv";
            // Filtre sur les fichiers
            openFileDialog1.Filter = "Tous les fichiers (*.csv)|*.csv";
            openFileDialog1.Multiselect = false;
            openFileDialog1.FilterIndex = 1;
            // Ouverture boite de dialogue OpenFile
            if (openFileDialog1.ShowDialog(this) == DialogResult.OK)
            {
                file = openFileDialog1.FileName;
                txt_Emplacement.Text = file;
                StreamReader sr = new StreamReader(file);
                // for set encoding
     //           string valFinal = null;
     //           string nomF = null;
                string strline = "";
                string[] _values = null;
      //          string[] noms = null;
      //          string[] noms1 = null;
                int x = 0;
                while (!sr.EndOfStream)
                {
                    x++;
                    strline = sr.ReadLine();
                    _values = strline.Split(';');

                            dataGridView1.Rows.Add(_values);
                    }

                sr.Close();
            }
      
        }

        private void checkBox1_Click(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
            {
                groupBox1.Visible = true;
            }
            else
            {
                groupBox1.Visible = false;
            }
        }

        private void Account_Plan_about_GL_File_Load(object sender, EventArgs e)
        {

        }

    }
}
