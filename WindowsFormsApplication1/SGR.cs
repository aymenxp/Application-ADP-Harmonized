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
    public partial class SGR : Form
    {
        public SGR()
        {
            InitializeComponent();
        }
        public int idSoc;
        public string codeSoc = string.Empty;
        public string NomSoc = string.Empty;
        TraitementGRFile tr = new TraitementGRFile();
        FormatDate form = new FormatDate();
        ServiceDAO service = new ServiceDAO();

        private void btn_Extract_Click(object sender, EventArgs e)
        {
            service.Delete_SGLTAB();
            string Date1 = form.LA_DATE(dateTimePicker1.Value.ToString());
            string Date2 = form.LA_DATE(dateTimePicker2.Value.ToString());
            string DateJour = form.LA_DATE(DateTime.Today.ToString());
            int MoisPe = int.Parse(dateTimePicker1.Value.ToString().Substring(3, 2));
            int AnneePe = int.Parse(dateTimePicker1.Value.ToString().Substring(6, 4));
            if (cmb_sens.Text == "Département")
            {
                tr.ExtraireDataSGL("P_SGR_", "Reconciliationledger.xls", Date1, Date2, DateJour, codeSoc, NomSoc, idSoc, MoisPe, AnneePe, "Dept");
            }
            if (cmb_sens.Text == "Service")
            {
                tr.ExtraireDataSGL("P_SGR_", "Reconciliationledger.xls", Date1, Date2, DateJour, codeSoc, NomSoc, idSoc, MoisPe, AnneePe, "Sce");
            }
            if (cmb_sens.Text == "Centre cout")
            {
                tr.ExtraireDataSGL("P_SGR_", "Reconciliationledger.xls", Date1, Date2, DateJour, codeSoc, NomSoc, idSoc, MoisPe, AnneePe, "CentreC");
            }

        }
    }
}
