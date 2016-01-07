using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using HarmonizedApp.Manager;
using System.Configuration;
using System.Data.OleDb;
using System.IO;

namespace HarmonizedApp
{
    public partial class Import_Rubrique : Form
    {

        ExtractionSRF ext = new ExtractionSRF();
        ServiceDAO serv = new ServiceDAO();

        public Import_Rubrique()
        {
            InitializeComponent();
        }

        private void btn_Emplacement_Click(object sender, EventArgs e)
        {
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
                string strline = "";
                string[] _values = null;
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

        private void btn_Import_Click(object sender, EventArgs e)
        {
            string CodeRub = string.Empty;
            string CodeCpt = string.Empty;
            string Desc = string.Empty;
            string DescCodeR = string.Empty;
            string Sens = string.Empty;
            bool rt = false;
            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                string code = dataGridView1.Rows[i].Cells[0].Value.ToString();
                string LibCodRub = dataGridView1.Rows[i].Cells[1].Value.ToString();
                int idSoc = int.Parse(dataGridView1.Rows[i].Cells[2].Value.ToString());
                int OrdreSeq = int.Parse(dataGridView1.Rows[i].Cells[3].Value.ToString());
                rt = rt = serv.insertRubrique(code, LibCodRub, idSoc.ToString(), OrdreSeq);
            }
            if (rt)
            {
                MessageBox.Show("Insert realized with Success");
            }
            else
            {
                MessageBox.Show("Insert Error", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

    }
}
