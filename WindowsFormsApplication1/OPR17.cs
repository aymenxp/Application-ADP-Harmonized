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
    public partial class OPR17 : Form
    {

        public int idSoc;
        public string codeSoc = string.Empty;
        public string NomSoc = string.Empty;
        TraitementGRFile tr = new TraitementGRFile();
        FormatDate form = new FormatDate();
        ServiceDAO service = new ServiceDAO();

        public OPR17()
        {
            InitializeComponent();
        }

        private void btn_Extract_Click(object sender, EventArgs e)
        {
            string Date1 = form.LA_DATE(dateTimePicker1.Value.ToString());
            int MoisPe1 = int.Parse(dateTimePicker1.Value.ToString().Substring(3, 2));
            int AnneePe1 = int.Parse(dateTimePicker1.Value.ToString().Substring(6, 4));

            tr.ExtraireDataOPR17("OPR17", ".xls", Date1, codeSoc, NomSoc, idSoc, MoisPe1, AnneePe1);
        }

        private void OPR17_Load(object sender, EventArgs e)
        {

        }
    }
}
