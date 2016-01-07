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
    public partial class Payment_Report : Form
    {

        public string CodeSoc;
        public int RefSoc;
        public string NomSoc;
        ExtractionSRF ext = new ExtractionSRF();
        ServiceDAO serv = new ServiceDAO();


        public Payment_Report()
        {
            InitializeComponent();
        }

        private void Payment_Report_Load(object sender, EventArgs e)
        {
            label7.Enabled = false;
            dateTimePicker1.Enabled = false;
            btn_Delete.Enabled = false;      
        }

        private void btn_Emplacement_Click(object sender, EventArgs e)
        {
            txt_Emplacement.Text = ext.EmplacementFich();
        }

        private void btn_Import_Click(object sender, EventArgs e)
        {
            string emplacement = txt_Emplacement.Text;
            ext.EcrireDataReport(RefSoc, emplacement, progressBar1);

        }

        private void button1_Click(object sender, EventArgs e)
        {
            int mois = int.Parse(dateTimePicker1.Value.Month.ToString());
            int annee = int.Parse(dateTimePicker1.Value.Year.ToString());
            if (MessageBox.Show("Would you like to delete the different DATA in the month : " + mois + " " + annee, "Attention", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
            {
                if (serv.Delete_PayReport(RefSoc,mois,annee) == true)
                {
                    MessageBox.Show("Delete realized with success");
                }
                else
                {
                    MessageBox.Show(" Error while delete this user", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void chk_delete_CheckedChanged(object sender, EventArgs e)
        {
            if (chk_delete.Checked)
            {
                label7.Enabled = true;
                dateTimePicker1.Enabled = true;
                btn_Delete.Enabled = true;               
            }
            else
            {
                label7.Enabled = false;
                dateTimePicker1.Enabled = false;
                btn_Delete.Enabled = false;
            }
        }
    }
}
