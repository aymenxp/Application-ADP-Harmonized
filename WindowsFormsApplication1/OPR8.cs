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
    public partial class OPR8 : Form
    {

        public int idSoc;
        public string codeSoc = string.Empty;
        public string NomSoc = string.Empty;
        TraitementGRFile tr = new TraitementGRFile();
        FormatDate form = new FormatDate();
        ServiceDAO service = new ServiceDAO();

        public OPR8()
        {
            InitializeComponent();
        }

        private void OPR8_Load(object sender, EventArgs e)
        {

        }

        private void btn_Extract_Click(object sender, EventArgs e)
        {
            string Date1 = form.LA_DATE(dateTimePicker1.Value.ToString());
            string Date2 = form.LA_DATE(dateTimePicker2.Value.ToString());
            int MoisPe1 = int.Parse(dateTimePicker1.Value.ToString().Substring(3, 2));
            int AnneePe1 = int.Parse(dateTimePicker1.Value.ToString().Substring(6, 4));
            int MoisPe2 = int.Parse(dateTimePicker2.Value.ToString().Substring(3, 2));
            int AnneePe2 = int.Parse(dateTimePicker2.Value.ToString().Substring(6, 4));

            tr.ExtraireDataOPR8("OPR8", ".xls", Date1, Date2, codeSoc, NomSoc, idSoc, MoisPe1, AnneePe1,MoisPe2,AnneePe2);
        }
    }
}
