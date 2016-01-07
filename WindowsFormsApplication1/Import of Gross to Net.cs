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
    public partial class Import_of_Gross_to_Net : Form
    {

        public int RefSoc;
        public string CodeSoc;
        public string NomSoc;


        ExtractionSRF ext = new ExtractionSRF();
        ServiceDAO serv = new ServiceDAO();

        public Import_of_Gross_to_Net()
        {
            InitializeComponent();
        }

        private void btn_Emplacement_Click(object sender, EventArgs e)
        {
            txt_Emplacement.Text = ext.EmplacementFich();
        }

        private void Import_of_Gross_to_Net_Load(object sender, EventArgs e)
        {

        }

        private void btn_Import_Click(object sender, EventArgs e)
        {
            string emplacement= txt_Emplacement.Text;
            if (RefSoc == 2)
            {
                ext.EcrireDataA001(RefSoc, emplacement, progressBar1);
            }
            if (RefSoc == 5)
            {
                ext.EcrireDataA002(RefSoc, emplacement, progressBar1);
            }
            if (RefSoc == 8)
            {
                ext.EcrireDataA003(RefSoc, emplacement, progressBar1);
            }

            if (RefSoc == 7)
            {
                ext.EcrireDataA004(RefSoc, emplacement, progressBar1);
            }
            // Sony Ericson
            if (RefSoc == 14)
            {
                ext.EcrireDataA021(RefSoc, emplacement, progressBar1);
            }

            if (RefSoc == 9)
            {
                ext.EcrireDataA010(RefSoc, emplacement, progressBar1);
            }

            if (RefSoc == 10)
            {
                ext.EcrireDataA011(RefSoc, emplacement, progressBar1);
            }
            if (RefSoc == 3)
            {
                ext.EcrireDataA014(RefSoc, emplacement, progressBar1);
            }
            if (RefSoc == 12)
            {
                ext.EcrireDataA015(RefSoc, emplacement, progressBar1);
            }
            if (RefSoc == 11)
            {
                ext.EcrireDataA013(RefSoc, emplacement, progressBar1);
            }
            if (RefSoc == 4)
            {
                ext.EcrireDataA020(RefSoc, emplacement, progressBar1);
            }

            if (RefSoc ==13)
            {
                ext.EcrireDataA016(RefSoc, emplacement, progressBar1);
            }
  
        }

        private void chk_delete_CheckedChanged(object sender, EventArgs e)
        {
            if (chk_delete.Checked)
            {
                groupBox1.Visible = true;
            }
            else
            {
                groupBox1.Visible = false;
            }
        }

        private void btn_delete_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Would you like to delete the Data of : " + serv.NomDuMois(dateTimePicker1.Value.Month), "Attention", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
            {
                int mois = dateTimePicker1.Value.Month;
                int annee = dateTimePicker1.Value.Year;
                if (RefSoc == 2)
                {
                    if (serv.Delete_GrossToNet("GROSSTONETA001",RefSoc,mois,annee) == true)
                    {
                        MessageBox.Show("Delete realised with success");
                    }
                    else
                    {
                        MessageBox.Show(" Error while delete this rubric ", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                if (RefSoc == 5)
                {
                    if (serv.Delete_GrossToNet("GROSSTONETA002", RefSoc, mois, annee) == true)
                    {
                        MessageBox.Show("Delete realised with success");
                    }
                    else
                    {
                        MessageBox.Show(" Error while delete this rubric ", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                if (RefSoc == 8)
                {
                    if (serv.Delete_GrossToNet("GROSSTONETA003", RefSoc, mois, annee) == true)
                    {
                        MessageBox.Show("Delete realised with success");
                    }
                    else
                    {
                        MessageBox.Show(" Error while delete this rubric ", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }

                if (RefSoc == 7)
                {
                    if (serv.Delete_GrossToNet("GROSSTONETA004", RefSoc, mois, annee) == true)
                    {
                        MessageBox.Show("Delete realised with success");
                    }
                    else
                    {
                        MessageBox.Show(" Error while delete this rubric ", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }

                if (RefSoc == 9)
                {
                    if (serv.Delete_GrossToNet("GROSSTONETA010", RefSoc, mois, annee) == true)
                    {
                        MessageBox.Show("Delete realised with success");
                    }
                    else
                    {
                        MessageBox.Show(" Error while delete this rubric ", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }

                if (RefSoc == 10)
                {
                    if (serv.Delete_GrossToNet("GROSSTONETA011", RefSoc, mois, annee) == true)
                    {
                        MessageBox.Show("Delete realised with success");
                    }
                    else
                    {
                        MessageBox.Show(" Error while delete this rubric ", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }

                if (RefSoc == 3)
                {
                    if (serv.Delete_GrossToNet("GROSSTONETA014", RefSoc, mois, annee) == true)
                    {
                        MessageBox.Show("Delete realised with success");
                    }
                    else
                    {
                        MessageBox.Show(" Error while delete this rubric ", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }

                if (RefSoc == 11)
                {
                    if (serv.Delete_GrossToNet("GROSSTONETA013", RefSoc, mois, annee) == true)
                    {
                        MessageBox.Show("Delete realised with success");
                    }
                    else
                    {
                        MessageBox.Show(" Error while delete this rubric ", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }

                if (RefSoc == 12)
                {
                    if (serv.Delete_GrossToNet("GROSSTONETA015", RefSoc, mois, annee) == true)
                    {
                        MessageBox.Show("Delete realised with success");
                    }
                    else
                    {
                        MessageBox.Show(" Error while delete this rubric ", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                if (RefSoc == 4)
                {
                    if (serv.Delete_GrossToNet("GROSSTONETA020", RefSoc, mois, annee) == true)
                    {
                        MessageBox.Show("Delete realised with success");
                    }
                    else
                    {
                        MessageBox.Show(" Error while delete this rubric ", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }

                if (RefSoc == 13)
                {
                    if (serv.Delete_GrossToNet("GROSSTONETA016", RefSoc, mois, annee) == true)
                    {
                        MessageBox.Show("Delete realised with success");
                    }
                    else
                    {
                        MessageBox.Show(" Error while delete this rubric ", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }

                if (RefSoc ==14)
                {
                    if (serv.Delete_GrossToNet("GROSSTONETA021", RefSoc, mois, annee) == true)
                    {
                        MessageBox.Show("Delete realised with success");
                    }
                    else
                    {
                        MessageBox.Show(" Error while delete this rubric ", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }



                
            }
        }
    }
}
