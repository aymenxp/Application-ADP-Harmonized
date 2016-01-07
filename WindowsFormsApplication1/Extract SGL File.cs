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
    public partial class Extract_SGL_File : Form
    {
        public Extract_SGL_File()
        {
            InitializeComponent();
        }

        public int idSoc;
        public string codeSoc = string.Empty;
        public string NomSoc = string.Empty;
        TraitementGLFILE tr = new TraitementGLFILE();
        FormatDate form = new FormatDate();

        private void Extract_SGL_File_Load(object sender, EventArgs e)
        {

        }

        private void btn_Extract_Click(object sender, EventArgs e)
        {
            string Date1 = form.LA_DATE(dateTimePicker1.Value.ToString());
            string Date2 = form.LA_DATE(dateTimePicker2.Value.ToString());
            string DateJour = form.LA_DATE(DateTime.Today.ToString());
            int MoisPe = int.Parse(dateTimePicker1.Value.ToString().Substring(3, 2));
            int AnneePe = int.Parse(dateTimePicker1.Value.ToString().Substring(6, 4));
            tr.ExtraireDataSGL("P_SGL_", "Summarygeneralledger.xls", Date1, Date2, DateJour, codeSoc, NomSoc, idSoc, MoisPe, AnneePe, "CentreCout");
        }
    }
}
