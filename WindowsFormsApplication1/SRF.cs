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
using System.Configuration;

namespace HarmonizedApp
{
    public partial class SRF : Form
    {

        ServiceDAO serv = new ServiceDAO();
        ExtractionSRF extra = new ExtractionSRF();
        FormatDate form = new FormatDate();
        private string connection_string = string.Empty;
        public int idSoc;
        public string NomSoc = string.Empty;
        public string CodeSoc = string.Empty;

        public SRF()
        {
            InitializeComponent();
        }

        private void SRF_Load(object sender, EventArgs e)
        {
            progressBar1.Visible = false; 
        }

        private void btn_Emplacement_Click(object sender, EventArgs e)
        {

        }

        private void btn_Extract_Click(object sender, EventArgs e)
        {
                bool PayNormal = radioButton1.Checked;
                bool Bonus = radioButton2.Checked;
                string Date1 = form.LA_DATE(dateTimePicker1.Value.ToString());
                string Date2 = form.LA_DATE(dateTimePicker2.Value.ToString());
                string DatePe = dateTimePicker3.Value.ToString().Substring(0, 10);
                string MoisPe = form.ExtraireMoisPe(dateTimePicker1.Value.ToString());
                string DateDeb = dateTimePicker1.Value.ToString().Substring(0, 10);
                string DateFin = dateTimePicker2.Value.ToString().Substring(0, 10);
                string Annee = dateTimePicker1.Value.ToString().Substring(6,4);
                extra.ExtraireDataSRF(Date1, Date2, CodeSoc, MoisPe, NomSoc, DateDeb, DateFin, PayNormal, Bonus, DatePe, idSoc,progressBar1,Annee);
            
        }

        private void btn_Emp2_Click(object sender, EventArgs e)
        {
        }

        private void btn_Import_Click(object sender, EventArgs e)
        {
            string file = string.Empty;
            int ligne = 0;
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Title = "Choisir le fichier à importer";
            // Extension par défaut			
            openFileDialog1.DefaultExt = "txt";
            // Filtre sur les fichiers
            openFileDialog1.Filter = "Tous les fichiers (*.txt)|*.txt";
            openFileDialog1.Multiselect = false;
            openFileDialog1.FilterIndex = 1;
            // Ouverture boite de dialogue OpenFile
            if (openFileDialog1.ShowDialog(this) == DialogResult.OK)
            {
                file = openFileDialog1.FileName;
                StreamReader sr = new StreamReader(file);
                string line = null;
                while ((line = sr.ReadLine()) != null)
                {
                    ligne = ligne + 1;

                    if (!(line.Equals("")))
                    {
                        string matricule = line.Substring(0, 10).Trim();
                        string nom = line.Substring(10, 80).Trim().Replace("'", "''");
                        string prenom = line.Substring(90, 20).Trim().Replace("'", "''");

                    }
                }
            }
        }
    }
}
