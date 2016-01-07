using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Configuration;
using System.Data.SqlClient;
using HarmonizedApp.Manager;

namespace HarmonizedApp
{
    public partial class Code_Memo : Form
    {
        public int idSoc;

        private string connection_string = string.Empty;
        ServiceDAO serv = new ServiceDAO();

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
        public Code_Memo()
        {
            InitializeComponent();
             if (connection_string != null)
            {
                cnx = new SqlConnection(connection_string);
            }
        }

        private void Code_Memo_Load(object sender, EventArgs e)
        {
            if (idSoc == 2)
            {
                connection_string = "Data Source="+serveur+";Initial Catalog="+ALCATEL+";Integrated Security=True";
                SqlConnection con = null;
                String ss = "SELECT Memo,CodeRubrique,Intitule FROM T_RUB";
                con = con = new SqlConnection(connection_string);
                con.Open();
                SqlCommand cmdx = new SqlCommand(ss, con);
                SqlDataReader dr = cmdx.ExecuteReader();
                System.Data.DataTable dt = new System.Data.DataTable();
                dt.Load(dr);

                foreach (System.Data.DataRow row in dt.Rows)
                {
                    dataGridView1.Rows.Add(row["Memo"].ToString(), row["CodeRubrique"].ToString(),
                    row["Intitule"].ToString());

                }
                
            }
            if (idSoc == 5)
            {
                connection_string = "Data Source=" + serveur + ";Initial Catalog=" + SPB + ";Integrated Security=True";
                SqlConnection con = null;
                String ss = "SELECT Memo,CodeRubrique,Intitule FROM T_RUB";
                con = con = new SqlConnection(connection_string);
                con.Open();
                SqlCommand cmdx = new SqlCommand(ss, con);
                SqlDataReader dr = cmdx.ExecuteReader();
                System.Data.DataTable dt = new System.Data.DataTable();
                dt.Load(dr);

                foreach (System.Data.DataRow row in dt.Rows)
                {
                    dataGridView1.Rows.Add(row["Memo"].ToString(), row["CodeRubrique"].ToString(),
                       row["Intitule"].ToString());
                }
                
                
            }
            if (idSoc == 8)
            {
                connection_string = "Data Source=" + serveur + ";Initial Catalog=" + COATS03 + ";Integrated Security=True";
                SqlConnection con = null;
                String ss = "SELECT Memo,CodeRubrique,Intitule FROM T_RUB";
                con = con = new SqlConnection(connection_string);
                con.Open();
                SqlCommand cmdx = new SqlCommand(ss, con);
                SqlDataReader dr = cmdx.ExecuteReader();
                System.Data.DataTable dt = new System.Data.DataTable();
                dt.Load(dr);

                foreach (System.Data.DataRow row in dt.Rows)
                {
                    dataGridView1.Rows.Add(row["Memo"].ToString(), row["CodeRubrique"].ToString(),
                    row["Intitule"].ToString());
                }
                
              
            }
            if (idSoc == 7)
            {
                connection_string = "Data Source=" + serveur + ";Initial Catalog=" + COATS04 + ";Integrated Security=True";
                SqlConnection con = null;
                String ss = "SELECT Memo,CodeRubrique,Intitule FROM T_RUB";
                con = con = new SqlConnection(connection_string);
                con.Open();
                SqlCommand cmdx = new SqlCommand(ss, con);
                SqlDataReader dr = cmdx.ExecuteReader();
                System.Data.DataTable dt = new System.Data.DataTable();
                dt.Load(dr);

                foreach (System.Data.DataRow row in dt.Rows)
                {
                    dataGridView1.Rows.Add(row["Memo"].ToString(), row["CodeRubrique"].ToString(),
                    row["Intitule"].ToString());
                }
                
              
            }
            if (idSoc == 9)
            {
                connection_string = "Data Source=" + serveur + ";Initial Catalog=" + STREAM + ";Integrated Security=True";
                SqlConnection con = null;
                String ss = "SELECT Memo,CodeRubrique,Intitule FROM T_RUB";
                con = con = new SqlConnection(connection_string);
                con.Open();
                SqlCommand cmdx = new SqlCommand(ss, con);
                SqlDataReader dr = cmdx.ExecuteReader();
                System.Data.DataTable dt = new System.Data.DataTable();
                dt.Load(dr);

                foreach (System.Data.DataRow row in dt.Rows)
                {
                    dataGridView1.Rows.Add(row["Memo"].ToString(), row["CodeRubrique"].ToString(),
                    row["Intitule"].ToString());
                }               
                
            }

            if (idSoc == 10)
            {
                connection_string = "Data Source=" + serveur + ";Initial Catalog=" + REDKNEE + ";Integrated Security=True";
                SqlConnection con = null;
                String ss = "SELECT Memo,CodeRubrique,Intitule FROM T_RUB";
                con = con = new SqlConnection(connection_string);
                con.Open();
                SqlCommand cmdx = new SqlCommand(ss, con);
                SqlDataReader dr = cmdx.ExecuteReader();
                System.Data.DataTable dt = new System.Data.DataTable();
                dt.Load(dr);

                foreach (System.Data.DataRow row in dt.Rows)
                {
                    dataGridView1.Rows.Add(row["Memo"].ToString(), row["CodeRubrique"].ToString(),
                    row["Intitule"].ToString());
                }
                
                
            }
            if (idSoc == 3)
            {
                connection_string = "Data Source=" + serveur + ";Initial Catalog=" + SCOGAT + ";Integrated Security=True";
                SqlConnection con = null;
                String ss = "SELECT Memo,CodeRubrique,Intitule FROM T_RUB";
                con = con = new SqlConnection(connection_string);
                con.Open();
                SqlCommand cmdx = new SqlCommand(ss, con);
                SqlDataReader dr = cmdx.ExecuteReader();
                System.Data.DataTable dt = new System.Data.DataTable();
                dt.Load(dr);

                foreach (System.Data.DataRow row in dt.Rows)
                {
                    dataGridView1.Rows.Add(row["Memo"].ToString(), row["CodeRubrique"].ToString(),
                    row["Intitule"].ToString());
                }
   
            }
            if (idSoc == 12)
            {
                connection_string = "Data Source=" + serveur + ";Initial Catalog=" + COCA + ";Integrated Security=True";
                SqlConnection con = null;
                String ss = "SELECT Memo,CodeRubrique,Intitule FROM T_RUB";
                con = con = new SqlConnection(connection_string);
                con.Open();
                SqlCommand cmdx = new SqlCommand(ss, con);
                SqlDataReader dr = cmdx.ExecuteReader();
                System.Data.DataTable dt = new System.Data.DataTable();
                dt.Load(dr);

                foreach (System.Data.DataRow row in dt.Rows)
                {
                    dataGridView1.Rows.Add(row["Memo"].ToString(), row["CodeRubrique"].ToString(),
                    row["Intitule"].ToString());
                }
       
            }
            if (idSoc == 11)
            {
                connection_string = "Data Source=" + serveur + ";Initial Catalog=" + N2SP + ";Integrated Security=True";
                SqlConnection con = null;
                String ss = "SELECT Memo,CodeRubrique,Intitule FROM T_RUB";
                con = con = new SqlConnection(connection_string);
                con.Open();
                SqlCommand cmdx = new SqlCommand(ss, con);
                SqlDataReader dr = cmdx.ExecuteReader();
                System.Data.DataTable dt = new System.Data.DataTable();
                dt.Load(dr);

                foreach (System.Data.DataRow row in dt.Rows)
                {
                    dataGridView1.Rows.Add(row["Memo"].ToString(), row["CodeRubrique"].ToString(),
                    row["Intitule"].ToString());
                }               
               
            }
            if (idSoc == 4)
            {
                connection_string = "Data Source=" + serveur + ";Initial Catalog=" + TTPC + ";Integrated Security=True";
                SqlConnection con = null;
                String ss = "SELECT Memo,CodeRubrique,Intitule FROM T_RUB";
                con = con = new SqlConnection(connection_string);
                con.Open();
                SqlCommand cmdx = new SqlCommand(ss, con);
                SqlDataReader dr = cmdx.ExecuteReader();
                System.Data.DataTable dt = new System.Data.DataTable();
                dt.Load(dr);

                foreach (System.Data.DataRow row in dt.Rows)
                {
                    dataGridView1.Rows.Add(row["Memo"].ToString(), row["CodeRubrique"].ToString(),
                    row["Intitule"].ToString());
                }                
                
            }

        }
    }
}
