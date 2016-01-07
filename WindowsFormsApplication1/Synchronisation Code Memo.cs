using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Configuration;
using System.Data.OleDb;
using HarmonizedApp.Manager;

namespace HarmonizedApp
{
    public partial class Synchronisation_Code_Memo : Form
    {

        private string connection_string = string.Empty;
        ServiceDAO service = new ServiceDAO();
        public int idSoc;

        private SqlConnection cnx;

        public SqlConnection Cnx
        {
            get { return cnx; }
            set { cnx = value; }
        }

        string serveur = ConfigurationManager.AppSettings["server"];
        string ALCATEL = ConfigurationManager.AppSettings["ALCATEL"];
        string SPB = ConfigurationManager.AppSettings["SPB"];
        string COATS03 = ConfigurationManager.AppSettings["COATS03"];
        string COATS04 = ConfigurationManager.AppSettings["COATS04"];
        string STREAM = ConfigurationManager.AppSettings["STREAM"];
        string REDKNEE = ConfigurationManager.AppSettings["REDKNEE"];
        string N2SP = ConfigurationManager.AppSettings["N2SP"];
        string SCOGAT = ConfigurationManager.AppSettings["SCOGAT"];
        string COCA = ConfigurationManager.AppSettings["COCA"];
        string TTPC = ConfigurationManager.AppSettings["TTPC"];
        string BG = ConfigurationManager.AppSettings["BG"];

        string A001 = ConfigurationManager.AppSettings["A001"];
        string A002 = ConfigurationManager.AppSettings["A002"];
        string A003 = ConfigurationManager.AppSettings["A003"];
        string A004 = ConfigurationManager.AppSettings["A004"];
        string A010 = ConfigurationManager.AppSettings["A010"];
        string A011 = ConfigurationManager.AppSettings["A011"];
        string A013 = ConfigurationManager.AppSettings["A013"];
        string A014 = ConfigurationManager.AppSettings["A014"];
        string A015 = ConfigurationManager.AppSettings["A015"];
        string A020 = ConfigurationManager.AppSettings["A020"];
        string A016 = ConfigurationManager.AppSettings["A016"];

        public Synchronisation_Code_Memo()
        {
            InitializeComponent();
        }

        private void Synchronisation_Code_Memo_Load(object sender, EventArgs e)
        {
            int Mois = int.Parse(dateTimePicker3.Value.ToString().Substring(3,2));
            int Annee=int.Parse(dateTimePicker3.Value.ToString().Substring(6,4));
            if (idSoc == 2)
            {
                ConnexionBaseSql(serveur, ALCATEL);
                RemplirGrid(A001, Mois, Annee);
            }
            if (idSoc == 5)
            {
                ConnexionBaseSql(serveur, SPB);
                RemplirGrid(A002, Mois, Annee);
            }
            if (idSoc == 8)
            {
                ConnexionBaseSql(serveur, COATS03);
                RemplirGrid(A003, Mois, Annee);
            }
            if (idSoc == 7)
            {
                ConnexionBaseSql(serveur, COATS04);
                RemplirGrid(A004, Mois, Annee);
            }
            if (idSoc == 9)
            {
                ConnexionBaseSql(serveur, STREAM);
                RemplirGrid(A010, Mois, Annee);
            }
            if (idSoc == 10)
            {
                ConnexionBaseSql(serveur, REDKNEE);
                RemplirGrid(A011, Mois, Annee);
            }
            if (idSoc == 3)
            {
                ConnexionBaseSql(serveur, SCOGAT);
                RemplirGrid(A014, Mois, Annee);
            }
            if (idSoc == 12)
            {
                ConnexionBaseSql(serveur, COCA);
                RemplirGrid(A015, Mois, Annee);
            }
            if (idSoc == 11)
            {
                ConnexionBaseSql(serveur, N2SP);
                RemplirGrid(A013, Mois, Annee);
            }
            if (idSoc == 4)
            {
                ConnexionBaseSql(serveur, TTPC);
                RemplirGrid(A020, Mois, Annee);
            }

            if (idSoc == 13)
            {
                ConnexionBaseSql(serveur, BG);
                RemplirGrid(A016, Mois, Annee);
            }
        }

        public void RemplirGrid(string NomTab,int Mois,int Annee)
        {
            if ((NomTab == "GROSSTONETA020") || (NomTab == "GROSSTONETA014") || (NomTab == "GROSSTONETA013") || (NomTab == "GROSSTONETA010"))
            {
            OleDbDataReader dr = service.RecupererEmp(NomTab, Mois, Annee,"IdOracle");
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Load(dr);
            
                foreach (System.Data.DataRow row in dt.Rows)
                {
                    dataGridView2.Rows.Add(row["IdOracle"].ToString(), row["Nom"].ToString(),
                    row["Prenom"].ToString());

                }
            }
            else
            {
                OleDbDataReader dr = service.RecupererEmp(NomTab, Mois, Annee, "Matricule");
                System.Data.DataTable dt = new System.Data.DataTable();
                dt.Load(dr);
                foreach (System.Data.DataRow row in dt.Rows)
                {
                    dataGridView2.Rows.Add(row["Matricule"].ToString(), row["Nom"].ToString(),
                    row["Prenom"].ToString());

                }
            }

        }

        public void ConnexionBaseSql(string serveur, string BaseAccess)
        {
            connection_string = "Data Source=" + serveur + ";Initial Catalog=" + BaseAccess + ";Integrated Security=True";
            SqlConnection con = null;
            String ss = "SELECT Memo,CodeRubrique,Intitule FROM T_RUB order by Memo ASC";
            con = con = new SqlConnection(connection_string);
            con.Open();
            SqlCommand cmdx = new SqlCommand(ss, con);
            SqlDataReader dr = cmdx.ExecuteReader();
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Load(dr);

            foreach (System.Data.DataRow row in dt.Rows)
            {
                string CodeM = row["Memo"].ToString();
                if ((CodeM != "NA")&&(CodeM !=""))
                {
                    dataGridView1.Rows.Add(CodeM, row["CodeRubrique"].ToString(),
                   row["Intitule"].ToString());
                }
            }

        }

        private void btn_Synchro_Click(object sender, EventArgs e)
        {
            progressBar1.Visible = true;
            progressBar1.Minimum = 0;
            progressBar1.Maximum = dataGridView2.RowCount;
            int Mois = int.Parse(dateTimePicker3.Value.ToString().Substring(3,2));
            int Annee=int.Parse(dateTimePicker3.Value.ToString().Substring(6,4));
            string CodeF = string.Empty;
            double ValeurI = 0.0;
            service.ViderTabCodeMemo();
            for (int j = 0; j < dataGridView2.RowCount; j++)
            {
                string Mat = dataGridView2.Rows[j].Cells[0].Value.ToString();
                progressBar1.Value = progressBar1.Value + 1;
                for (int i = 0; i < dataGridView1.RowCount; i++)
                {
                    if (i <= dataGridView1.RowCount)
                    {
                            string CodeM = dataGridView1.Rows[i].Cells[0].Value.ToString();
           //               int ligne = VerifFinGrid(i);
                            try
                            {
                                CodeF = dataGridView1.Rows[i + 1].Cells[0].Value.ToString();
                            }
                            catch
                            {
                                CodeM = dataGridView1.Rows[i].Cells[0].Value.ToString();
                            }

                            string CodeRub = "A"+dataGridView1.Rows[i].Cells[1].Value.ToString();
                        
                            if (CodeM == CodeF)
                            {
                                double val = service.RecupererVF(idSoc, Mois, Annee, CodeRub, TypeSoc(idSoc), Mat);
                                ValeurI = ValeurI + val;
                                if (i == dataGridView1.RowCount-1)
                                {
                                    service.insertPaymentReport(CodeM, ValeurI.ToString("#.###"), Mat, idSoc.ToString());
                                    ValeurI = 0;
                                }
                            }
                            else
                            {
                                double val = service.RecupererVF(idSoc, Mois, Annee, CodeRub, TypeSoc(idSoc), Mat);
                                ValeurI = ValeurI + val;
                                service.insertPaymentReport(CodeM, ValeurI.ToString("#.###"), Mat, idSoc.ToString());
                                ValeurI = 0;
                            }

                    }
                }
            }
            AffectationCodeMemo();
            progressBar1.Minimum = 0;
            progressBar1.Visible = false;

            MessageBox.Show("Synchronisation realized with Success");

        }

        public void AffectationCodeMemo()
        {
            double Val1 = 0.0;
            double Val2 = 0.0;
            int Mois = int.Parse(dateTimePicker3.Value.ToString().Substring(3,2));
            int Annee=int.Parse(dateTimePicker3.Value.ToString().Substring(6,4));
        for(int i=0;i<dataGridView2.RowCount;i++)
        {
            string Mat = dataGridView2.Rows[i].Cells[0].Value.ToString();
            if (idSoc == 2)
            {
                Val1 = service.SUMRetenue("A61109", A001, Mois, Annee, Mat);
                Val2 = service.SUMRetenue("A61209", A001, Mois, Annee, Mat);
                double ValRetCNSS = Val1 + Val2;
                service.insertPaymentReport("400", ValRetCNSS.ToString("#.###"), Mat, idSoc.ToString());
                double ValAss1 = service.SUMRetenue("A62109", A001, Mois, Annee, Mat);
                double ValAss2 = service.SUMRetenue("A62309", A001, Mois, Annee, Mat);
                double ValAss = ValAss1 - ValAss2;
                service.insertPaymentReport("450", ValAss.ToString("#.###"), Mat, idSoc.ToString());
                             
            }
            if (idSoc == 5)
            {
                Val1 = service.SUMRetenue("A61109", A002, Mois, Annee, Mat);
                double ValRetCNSS = Val1;
                service.insertPaymentReport("400", ValRetCNSS.ToString("#.###"), Mat, idSoc.ToString());
                double ValAss1 = service.SUMRetenue("A71159", A002, Mois, Annee, Mat);
                double ValAss = ValAss1;
                service.insertPaymentReport("450", ValAss.ToString("#.###"), Mat, idSoc.ToString());
              
            }
            if (idSoc == 8)
            {
                Val1 = service.SUMRetenue("A61109", A003, Mois, Annee, Mat);
                Val2 = service.SUMRetenue("A61209", A003, Mois, Annee, Mat);
                double ValRetCNSS = Val1 + Val2;
                service.insertPaymentReport("400", ValRetCNSS.ToString("#.###"), Mat, idSoc.ToString());
                double ValAss1 = service.SUMRetenue("A62109", A003, Mois, Annee, Mat);
                double ValAss = ValAss1;
                service.insertPaymentReport("450", ValAss.ToString("#.###"), Mat, idSoc.ToString());
              
            }
            if (idSoc == 7)
            {
                Val1 = service.SUMRetenue("A61109", A004, Mois, Annee, Mat);
                Val2 = service.SUMRetenue("A6120", A004, Mois, Annee, Mat);
                double ValRetCNSS = Val1 + Val2;
                service.insertPaymentReport("400", ValRetCNSS.ToString("#.###"), Mat, idSoc.ToString());
                double ValAss1 = service.SUMRetenue("A621099", A004, Mois, Annee, Mat);
                double ValAss = ValAss1;
                service.insertPaymentReport("450", ValAss.ToString("#.###"), Mat, idSoc.ToString());
            
            }
            if (idSoc == 9)
            {
                Val1 = service.SUMRetenue("A61109", A010, Mois, Annee, Mat);
                double ValRetCNSS = Val1;
                service.insertPaymentReport("400", ValRetCNSS.ToString("#.###"), Mat, idSoc.ToString());
                double ValAss1 = service.SUMRetenue("A71159", A010, Mois, Annee, Mat);
                double ValAss = ValAss1;
                service.insertPaymentReport("450", ValAss.ToString("#.###"), Mat, idSoc.ToString());
             
            }
            if (idSoc == 10)
            {
                Val1 = service.SUMRetenue("A61109", A011, Mois, Annee, Mat);
                double ValRetCNSS = Val1;
                service.insertPaymentReport("400", ValRetCNSS.ToString("#.###"), Mat, idSoc.ToString());            
              
            }
            if (idSoc == 3)
            {
                Val1 = service.SUMRetenue("A61109", A014, Mois, Annee, Mat);
                Val2 = service.SUMRetenue("A61209", A014, Mois, Annee, Mat);
                double ValRetCNSS = Val1 + Val2;
                service.insertPaymentReport("400", ValRetCNSS.ToString("#.###"), Mat, idSoc.ToString());
                double ValAss1 = service.SUMRetenue("A711599", A014, Mois, Annee, Mat);
                double ValAss2 = service.SUMRetenue("A711899", A014, Mois, Annee, Mat);
                double ValAss = ValAss1 - ValAss2;
                service.insertPaymentReport("450", ValAss.ToString("#.###"), Mat, idSoc.ToString());
             
            }
            if (idSoc == 12)
            {
                Val1 = service.SUMRetenue("A61109", A015, Mois, Annee, Mat);
                Val2 = service.SUMRetenue("A61209", A015, Mois, Annee, Mat);
                double ValRetCNSS = Val1 + Val2;
                service.insertPaymentReport("400", ValRetCNSS.ToString("#.###"), Mat, idSoc.ToString());
           
            }
            if (idSoc == 11)
            {
                Val1 = service.SUMRetenue("A61109", A013, Mois, Annee, Mat);
                double ValRetCNSS = Val1;
                service.insertPaymentReport("400", ValRetCNSS.ToString("#.###"), Mat, idSoc.ToString());
                double ValAss1 = service.SUMRetenue("A71159", A013, Mois, Annee, Mat);
                double ValAss = ValAss1;
                service.insertPaymentReport("450", ValAss.ToString("#.###"), Mat, idSoc.ToString());
            
            }
            if (idSoc == 4)
            {
                Val1 = service.SUMRetenue("A61109", A020, Mois, Annee, Mat);
                Val2 = service.SUMRetenue("A61209", A020, Mois, Annee, Mat);
                double ValRetCNSS = Val1 + Val2;
                service.insertPaymentReport("400", ValRetCNSS.ToString("#.###"), Mat, idSoc.ToString());
                double ValAss1 = service.SUMRetenue("A71159", A020, Mois, Annee, Mat);
                double ValAss2 = service.SUMRetenue("A62019", A020, Mois, Annee, Mat);
                double ValAss = ValAss1 - ValAss2;
                service.insertPaymentReport("450", ValAss.ToString("#.###"), Mat, idSoc.ToString());
            
            }
            }
        }

        public int VerifFinGrid(int ligne)
        {
            int ValF=0;
            if(ligne>=dataGridView1.RowCount)
            {
                ValF = 0;
            }
            else
            {
                ValF = ligne;
            }
            return ValF;
        }

        private void dateTimePicker3_ValueChanged(object sender, EventArgs e)
        {
            dataGridView2.Rows.Clear();
            int Mois = int.Parse(dateTimePicker3.Value.ToString().Substring(3, 2));
            int Annee = int.Parse(dateTimePicker3.Value.ToString().Substring(6, 4));
            if (idSoc == 2)
            {
                RemplirGrid(A001, Mois, Annee);
            }
            if (idSoc == 5)
            {
                RemplirGrid(A002, Mois, Annee);
            }
            if (idSoc == 8)
            {
                RemplirGrid(A003, Mois, Annee);
            }
            if (idSoc == 7)
            {
                RemplirGrid(A004, Mois, Annee);
            }
            if (idSoc == 9)
            {
                RemplirGrid(A010, Mois, Annee);
            }
            if (idSoc == 10)
            {
                RemplirGrid(A011, Mois, Annee);
            }
            if (idSoc == 3)
            {
                RemplirGrid(A014, Mois, Annee);
            }
            if (idSoc == 12)
            {
                RemplirGrid(A015, Mois, Annee);
            }
            if (idSoc == 11)
            {
                RemplirGrid(A013, Mois, Annee);
            }
            if (idSoc == 4)
            {
                RemplirGrid(A020, Mois, Annee);
            }
        }

        public string TypeSoc(int idSocc)
        {
            string TypeF = string.Empty;
            if (idSoc == 2)
            {
                TypeF = A001;
            }
            if (idSoc == 3)
            {
                TypeF = A014;
            }
            if (idSoc == 4)
            {
                TypeF = A020;
            }
            if (idSoc == 5)
            {
                TypeF = A002;
            }
            if (idSoc == 7)
            {
                TypeF = A004;
            }
            if (idSoc == 8)
            {
                TypeF = A003;
            }
            if (idSoc == 9)
            {
                TypeF = A010;
            }
            if (idSoc == 10)
            {
                TypeF = A011;
            }
            if (idSoc == 11)
            {
                TypeF = A013;
            }
            if (idSoc == 12)
            {
                TypeF = A015;
            }
            return TypeF;
        }

    }
}
