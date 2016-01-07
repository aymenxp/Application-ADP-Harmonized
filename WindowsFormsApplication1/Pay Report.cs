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
    public partial class Pay_Report : Form
    {
        public Pay_Report()
        {
            InitializeComponent();
        }

        public int idSoc;
        public string codeSoc = string.Empty;
        public string NomSoc = string.Empty;
        Extract_SEP sep = new Extract_SEP();
        FormatDate form = new FormatDate();

        private void Pay_Report_Load(object sender, EventArgs e)
        {

        }

        private void btn_SEP_Click(object sender, EventArgs e)
        {
            string Date1 = form.LA_DATE(dateTimePicker1.Value.ToString());
            string Date2 = form.LA_DATE(dateTimePicker2.Value.ToString());
            int MoisPe = int.Parse(dateTimePicker1.Value.ToString().Substring(3, 2));
            int annee = int.Parse(dateTimePicker1.Value.ToString().Substring(6,4));
            sep.ConstructSEP(Date1, Date2, codeSoc, NomSoc, progressBar1, idSoc, MoisPe,annee);
        }
    }
}
