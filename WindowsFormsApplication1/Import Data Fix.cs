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
    public partial class Import_Data_Fix : Form
    {

        public int RefSoc;
        public string CodeSoc;
        public string NomSoc;

        ExtractionSRF ext = new ExtractionSRF();
        ServiceDAO serv = new ServiceDAO();

        public Import_Data_Fix()
        {
            InitializeComponent();
        }

        private void btn_Emplacement_Click(object sender, EventArgs e)
        {
            txt_Emplacement.Text = ext.EmplacementFich();
        }

        private void btn_Import_Click(object sender, EventArgs e)
        {
            serv.Delete_DATAFIXE();
            bool verif = ext.EcrireDataFixe(RefSoc, txt_Emplacement.Text, progressBar1);
        }
    }
}
