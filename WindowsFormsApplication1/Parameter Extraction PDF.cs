using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using iTextSharp.text.pdf;
using iTextSharp.text;
using System.IO;
using ICSharpCode.SharpZipLib;
using ICSharpCode.SharpZipLib.Zip;
using System.Reflection;
using System.Threading;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Net.Mail;
using System.Net;
using HarmonizedApp.Manager;
using System.Configuration;

namespace HarmonizedApp
{
    public partial class Parameter_Extraction_PDF : Form
    {

        ServiceDAO serv = new ServiceDAO();
        string currpath;
        public string sourcePdfPath;
        private string pathFolder = string.Empty;
        private string idSociete = string.Empty;
        private string TransMode = string.Empty;
        private string Libelles = string.Empty;
        private string fileType = string.Empty;
        private string startDate = string.Empty;
        private string endDate = string.Empty;

        public Parameter_Extraction_PDF()
        {
            InitializeComponent();
            currpath = Directory.GetCurrentDirectory();
            discoInitialiseFrame();
        }

        private void discoInitialiseFrame()
        {
            pathFolder = ConfigurationManager.AppSettings["ExportFolder"];
            txt_Emplacement.Text = pathFolder;
        }

        private void Parameter_Extraction_PDF_Load(object sender, EventArgs e)
        {
            List<DataRaison> list = serv.GetSociteList();
            cmb_NomSoc.DataSource = list.ToArray();
            cmb_NomSoc.ValueMember = "societe";
            cmb_NomSoc.DisplayMember = "societe";
            if (list.Count > 0)
            {
                cmb_NomSoc.SelectedIndex = 0;
            }
            progressBar1.Visible = false;
            button3.Enabled = false;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Title = "Choisir le fichier à importer";
            // Extension par défaut			
            openFileDialog1.DefaultExt = "pdf";
            // Filtre sur les fichiers
            openFileDialog1.Filter = "Tous les fichiers (*.*)|*.*";
            openFileDialog1.Multiselect = false;
            openFileDialog1.FilterIndex = 1;
            if (openFileDialog1.ShowDialog(this) == DialogResult.OK) // Test result.
            {
                sourcePdfPath = openFileDialog1.FileName;
                this.textBox1.Text = sourcePdfPath;
                int taille = textBox1.Text.Length - textBox1.Text.Substring(0, textBox1.Text.LastIndexOf(".")).Length;
                if ((textBox1.Text.Substring(textBox1.Text.LastIndexOf("."), taille) == ".pdf") || (textBox1.Text.Substring(textBox1.Text.LastIndexOf("."), taille) == ".PDF"))
                {
                    button3.Enabled = true;
                }
                else
                {
                    button3.Enabled = false;
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog();

            if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
            {
                pathFolder = folderBrowserDialog.SelectedPath;
                txt_Emplacement.Text = pathFolder;
                /*           Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
                config.AppSettings.Settings.Remove("ExportFolder");
                config.AppSettings.Settings.Add("ExportFolder", pathFolder);
                config.Save(ConfigurationSaveMode.Modified);
                ConfigurationManager.RefreshSection("appSettings");  */
            }
        }

        public string LeMois(int Mois)
        {
            string Selstatus = string.Empty;
            switch (Mois)
            {
                case 1:
                    Selstatus = "Janvier";
                    break;
                case 2:
                    Selstatus = "Février";
                    break;
                case 3:
                    Selstatus = "Mars";
                    break;
                case 4:
                    Selstatus = "Avril";
                    break;
                case 5:
                    Selstatus = "Mai";
                    break;
                case 6:
                    Selstatus = "Juin";
                    break;
                case 7:
                    Selstatus = "Juill";
                    break;
                case 8:
                    Selstatus = "Aout";
                    break;
                case 9:
                    Selstatus = "Septembre";
                    break;
                case 10:
                    Selstatus = "Octobre";
                    break;
                case 11:
                    Selstatus = "Novembre";
                    break;
                case 12:
                    Selstatus = "Decembre";
                    break;
                default:
                    Selstatus = "";
                    break;
            }

            return Selstatus;
        }

        private void Dir(string directory)
        {
            string[] files = Directory.GetFiles(directory); // pour avoir les noms des fichiers et sous-répertoires
            for (int i = 0; i < files.Length; i++)
            {
                listBox1.Items.Add(files[i]);
            }
            listBox1.Refresh();
        }

        private void btn_Import_Click(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();
            int ligne = 0;
            string file = string.Empty;
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
                        dataGridView1.Rows.Add(matricule, nom, prenom);

                    }
                }
            }

            txt_NbreSal.Text = dataGridView1.RowCount.ToString();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (this.cmb_NomSoc.SelectedIndex != -1)
            {
                idSociete = ((DataRaison)this.cmb_NomSoc.SelectedItem).CodeSoc;
                progressBar1.Visible = true;
                string inputpath = textBox1.Text;
                PdfReader reader = new PdfReader(inputpath);
                int interval = 1;

                for (int pagenumber = 1; pagenumber <= reader.NumberOfPages; pagenumber += interval)
                {
                    SplitAndSaveInterval(textBox1.Text, idSociete, currpath, pagenumber, interval);
                    progressBar1.Value = progressBar1.Value + 1;      

                }
                MessageBox.Show("Operation completed successfully");
                Dir(txt_Emplacement.Text);
                btn_Compresser.Enabled = true;
                progressBar1.Visible = false;
                progressBar1.Value = 0;
            }
        }

        private int SplitAndSaveInterval(string inputPath, string CodeSoc, string outputPath, int startpage, int interval)
        {
            string chemain = null;
            int ligne = 0;
            int Compteur = int.Parse(txt_NbreSal.Text);
            FileInfo file = new FileInfo(inputPath);
            string name = file.Name.Substring(0, file.Name.LastIndexOf("."));
            string chaine = null; string matricule = null; string nom = null; string prenom = null;
            string mail = string.Empty;
            DataTable dm = serv.AffichageParam(idSociete);
            foreach (System.Data.DataRow reze in dm.Rows)
            {
                TransMode = reze["TransMode"].ToString();
                fileType = reze["FileType"].ToString();
                startDate = reze["StartDate"].ToString();
                endDate = reze["EndDate"].ToString();

            }
            string k = VerifierCodeEnvoie(idSociete);
            using (PdfReader reader = new PdfReader(inputPath))
            {

                for (int i = 0; i < dataGridView1.RowCount; i++)
                {
                    Document document = new Document();
                    ligne = ligne + 1;
                    matricule = dataGridView1.Rows[i].Cells[0].Value.ToString();
                    nom = dataGridView1.Rows[i].Cells[1].Value.ToString();
                    prenom = dataGridView1.Rows[i].Cells[2].Value.ToString();
                    chaine = TransMode + "_" + fileType + "_" + CodeSoc + "_" + startDate + "_" + endDate + "_00_V2_0000_00000_FILE_" + matricule + "_" + nom.Replace("�", "e") + " " + prenom.Replace("�", "e");
                    chemain = txt_Emplacement.Text + "\\" + chaine + ".pdf";
                    PdfCopy copy = new PdfCopy(document, new FileStream(chemain, FileMode.Create));
                    document.Open();
                    copy.AddPage(copy.GetImportedPage(reader, ligne));
                    document.Close();
                }

                return reader.NumberOfPages;
            }
        }

        private void btn_Compresser_Click(object sender, EventArgs e)
        {

        }

        public string VerifierCodeEnvoie(string idSociete)
        {
            int valFinal = 0;
            int valeur = serv.recapIncrement(idSociete.ToString());
            if (valeur == 0)
            {
                serv.insertIncrement(idSociete.ToString(), 11);
                valFinal = 11;
            }
            else
            {
                valFinal = valeur + 1;
                serv.modifIncrementation(idSociete, valFinal);
            }
            string increment = serv.Conc("n", 5, valFinal.ToString());
            return increment;
        }

        public String Conc(String type, int lentgh, String valeur)
         {
             int longueurManquante = lentgh - valeur.Length;
             if (longueurManquante > 0)
             {
                 if (type == "x")
                 {
                     for (int i = 0; i < longueurManquante; i++)
                     {
                         valeur = valeur + " ";
                     }
                 }
                 else
                 {
                     for (int i = 0; i < longueurManquante; i++)
                     {
                         valeur = "0" + valeur;
                     }
                 }

             }

             return valeur;

         }

    } 
}
