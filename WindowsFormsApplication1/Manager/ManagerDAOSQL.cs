using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Configuration;
using System.Data.SqlClient;
using System.Data;

namespace HarmonizedApp.Manager
{
   public class ManagerDAOSQL
    {
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

        public string LaBase(int refSoc)
        {
            string NomDeTab = string.Empty;
            if (refSoc == 2)
            {
                NomDeTab = ALCATEL;
            }
            if (refSoc == 5)
            {
                NomDeTab = SPB;
            }
            if (refSoc == 8)
            {
                NomDeTab = COATS03;
            }
            if (refSoc == 7)
            {
                NomDeTab = COATS04;
            }
            if (refSoc == 9)
            {
                NomDeTab = STREAM;
            }
            if (refSoc == 10)
            {
                NomDeTab = REDKNEE;
            }
            if (refSoc == 11)
            {
                NomDeTab = N2SP;
            }
            if (refSoc == 3)
            {
                NomDeTab = SCOGAT;
            }
            if (refSoc == 12)
            {
                NomDeTab = COCA;
            }
            if (refSoc == 4)
            {
                NomDeTab = TTPC;
            }

            return NomDeTab;

        }

        private string connection_string(int RefSoc)
        {
            string connection_string = string.Empty;
            connection_string = "Data Source=" + serveur + ";Initial Catalog=" + LaBase(RefSoc) + ";Integrated Security=True";
            return connection_string;
        }

        private SqlConnection cnx;

        public SqlConnection Cnx
        {
            get { return cnx; }
            set { cnx = value; }
        }

        public DataTable TestSens(string CodeRub,int RefSoc)
        {
            DataTable dt = new DataTable();
            SqlConnection con = null;
            try
            {
                con = new SqlConnection(connection_string(RefSoc));
                SqlDataAdapter adapt = new SqlDataAdapter("select distinct TypeRubrique,Gain from T_RUB where CodeRubrique=" + CodeRub + "", con);
                adapt.Fill(dt);

            }
            catch (Exception ex)
            {
                ex.ToString();
            }
            finally
            {
                if (con != null)
                {
                    con.Close();
                    con = null;
                }
            }

            return dt;

        }

    }
}
