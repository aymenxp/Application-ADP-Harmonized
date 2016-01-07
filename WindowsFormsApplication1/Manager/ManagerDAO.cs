using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.OleDb;
using System.Data;
using System.Configuration;
using System.Collections;
using System.Windows.Forms;

namespace HarmonizedApp.Manager
{
    class ManagerDAO
    {
        private string connection_string = @"Provider=Microsoft.Jet.OLEDB.4.0;" +
               @"Data source= " + Application.StartupPath + "\\ApplicationADP.mdb";

        private static ManagerDAO instance;

        private OleDbConnection cnx;

        public OleDbConnection conn = null;

        public OleDbConnection Cnx
        {
            get { return cnx; }
            set { cnx = value; }
        }

        public static ManagerDAO GetInstance()
        {
            if (instance == null)
            {
                instance = new ManagerDAO();
            }
            return instance;
        }

        private ManagerDAO()
        {
            if (connection_string != null)
            {
                cnx = new OleDbConnection(connection_string);
            }
        }


        public OleDbDataReader AffichageListeUsers()
        {
            OleDbConnection con = null;
            String ss = "select  * from USERS ORDER BY Login ASC";
            con = con = new OleDbConnection(connection_string);
            con.Open();
            OleDbCommand cmdx = new OleDbCommand(ss, con);
            OleDbDataReader dr = cmdx.ExecuteReader();
            return dr;
        }

        public OleDbDataReader AffichageListeRubrique(int idSoc)
        {
            OleDbConnection con = null;
            String ss = "select  * from ListeDesRubrique where idSoc='" + idSoc + "' ORDER BY OrdreSeq ASC";
            con = con = new OleDbConnection(connection_string);
            con.Open();
            OleDbCommand cmdx = new OleDbCommand(ss, con);
            OleDbDataReader dr = cmdx.ExecuteReader();
            return dr;
        }

        public OleDbDataReader AffichageListeGLFILE(int idSoc)
        {
            OleDbConnection con = null;
            String ss = "select  * from GLFILE where idSoc=" + idSoc + "";
            con = con = new OleDbConnection(connection_string);
            con.Open();
            OleDbCommand cmdx = new OleDbCommand(ss, con);
            OleDbDataReader dr = cmdx.ExecuteReader();
            return dr;
        }

        public OleDbDataReader AffichageListeCompany()
        {
            OleDbConnection con = null;
            String ss = "select  * from Societe ORDER BY NomSoc ASC";
            con = con = new OleDbConnection(connection_string);
            con.Open();
            OleDbCommand cmdx = new OleDbCommand(ss, con);
            OleDbDataReader dr = cmdx.ExecuteReader();
            return dr;
        }

        public bool Delete_User(int mat)
        {
            OleDbConnection con;
            bool etat = false;
            con = new OleDbConnection(connection_string);
            con.Open();
            string Delete_Soc = @"DELETE FROM USERS WHERE N° = " + mat + "";
            OleDbCommand Delete_T_Soc_Commande = new OleDbCommand(Delete_Soc, con);
            try
            {
                int count = Delete_T_Soc_Commande.ExecuteNonQuery();
                etat = true;
            }
            catch (OleDbException)
            {
                etat = false;
            }
            finally
            {
                con.Close();
            }

            return etat;

        }

        public bool ViderTabCodeMemo()
        {
            OleDbConnection con;
            bool etat = false;
            con = new OleDbConnection(connection_string);
            con.Open();
            string Delete_Soc = @"DELETE FROM CODEMEMO";
            OleDbCommand Delete_T_Soc_Commande = new OleDbCommand(Delete_Soc, con);
            try
            {
                int count = Delete_T_Soc_Commande.ExecuteNonQuery();
                etat = true;
            }
            catch (OleDbException)
            {
                etat = false;
            }
            finally
            {
                con.Close();
            }

            return etat;

        }

        public bool Delete_PayReport(int idSoc,int mois,int annee)
        {
            OleDbConnection con;
            bool etat = false;
            con = new OleDbConnection(connection_string);
            con.Open();
            string Delete_Soc = @"DELETE FROM PAYMENTREPORT WHERE RefSoc = " + idSoc + " and MoisPaie="+mois+" and AnneePaie="+annee+"";
            OleDbCommand Delete_T_Soc_Commande = new OleDbCommand(Delete_Soc, con);
            try
            {
                int count = Delete_T_Soc_Commande.ExecuteNonQuery();
                etat = true;
            }
            catch (OleDbException)
            {
                etat = false;
            }
            finally
            {
                con.Close();
            }

            return etat;

        }

        public bool Delete_Rubric(int mat)
        {
            OleDbConnection con;
            bool etat = false;
            con = new OleDbConnection(connection_string);
            con.Open();
            string Delete_Soc = @"DELETE FROM ListeDesRubrique WHERE Num = " + mat + "";
            OleDbCommand Delete_T_Soc_Commande = new OleDbCommand(Delete_Soc, con);
            try
            {
                int count = Delete_T_Soc_Commande.ExecuteNonQuery();
                etat = true;
            }
            catch (OleDbException)
            {
                etat = false;
            }
            finally
            {
                con.Close();
            }

            return etat;

        }

        public bool Delete_GLFILE(int mat)
        {
            OleDbConnection con;
            bool etat = false;
            con = new OleDbConnection(connection_string);
            con.Open();
            string Delete_Soc = @"DELETE FROM GLFILE WHERE Num = " + mat + "";
            OleDbCommand Delete_T_Soc_Commande = new OleDbCommand(Delete_Soc, con);
            try
            {
                int count = Delete_T_Soc_Commande.ExecuteNonQuery();
                etat = true;
            }
            catch (OleDbException)
            {
                etat = false;
            }
            finally
            {
                con.Close();
            }

            return etat;

        }

        public bool Delete_GrossToNet(string NomTab,int idSoc,int mois,int annee)
        {
            OleDbConnection con;
            bool etat = false;
            con = new OleDbConnection(connection_string);
            con.Open();
            string Delete_Soc = @"DELETE FROM " + NomTab + " WHERE IdSoc = " + idSoc + " and "
            + "Mois=" + mois + " and Annee="+annee+"";
            OleDbCommand Delete_T_Soc_Commande = new OleDbCommand(Delete_Soc, con);
            try
            {
                int count = Delete_T_Soc_Commande.ExecuteNonQuery();
                etat = true;
            }
            catch (OleDbException)
            {
                etat = false;
            }
            finally
            {
                con.Close();
            }

            return etat;

        }

        public bool Delete_Company(int mat)
        {
            OleDbConnection con;
            bool etat = false;
            con = new OleDbConnection(connection_string);
            con.Open();
            string Delete_Soc = @"DELETE FROM Societe WHERE Num = " + mat + "";
            OleDbCommand Delete_T_Soc_Commande = new OleDbCommand(Delete_Soc, con);
            try
            {
                int count = Delete_T_Soc_Commande.ExecuteNonQuery();
                etat = true;
            }
            catch (OleDbException)
            {
                etat = false;
            }
            finally
            {
                con.Close();
            }

            return etat;

        }

        public bool Delete_SGLTAB()
        {
            OleDbConnection con;
            bool etat = false;
            con = new OleDbConnection(connection_string);
            con.Open();
            string Delete_Soc = @"DELETE FROM SGLFINAL";
            OleDbCommand Delete_T_Soc_Commande = new OleDbCommand(Delete_Soc, con);
            try
            {
                int count = Delete_T_Soc_Commande.ExecuteNonQuery();
                etat = true;
            }
            catch (OleDbException)
            {
                etat = false;
            }
            finally
            {
                con.Close();
            }

            return etat;

        }

        public bool insertUserWKF(string login,string pass)
        {
            OleDbConnection con = null;
            bool etat = false;
            try
            {
                con = new OleDbConnection(connection_string);
                con.Open();
                string select = "insert into Users (LOGIN,PASS)"
               + " values('" + login + "','" + pass + "')";
                OleDbCommand cmd = new OleDbCommand(select, con);
                cmd.ExecuteNonQuery();
                etat = true;
            }

            catch (Exception ex)
            {
                etat = false;
            }
            finally
            {
                if (con != null)
                {
                    con.Close();
                    con = null;
                }
            }

            return etat;
        }

        public bool insertParamGlFile(string CodeRub,string LibCodRub,string CodeCpt,string Desc,string Sens,int IdSoc)
        {
            OleDbConnection con = null;
            bool etat = false;
            try
            {
                con = new OleDbConnection(connection_string);
                con.Open();
                string select = "insert into GLFILE(CodeRub,LibCodeRub,CodeCpt,Libelles,Sens,IdSoc)"
               + " values('" + CodeRub + "','" + LibCodRub + "','" + CodeCpt + "','" + Desc + "','" + Sens + "'," + IdSoc + ")";
                OleDbCommand cmd = new OleDbCommand(select, con);
                cmd.ExecuteNonQuery();
                etat = true;
            }

            catch (Exception ex)
            {
                etat = false;
            }
            finally
            {
                if (con != null)
                {
                    con.Close();
                    con = null;
                }
            }

            return etat;
        }

        public bool insertDATAGLFILE(string CodeRub,string GLACCOUNT, string GLDESCRIPTION, string COSTCENTER, string Desc, string Sens,string valeur, int IdSoc)
        {
            OleDbConnection con = null;
            bool etat = false;
            try
            {
                con = new OleDbConnection(connection_string);
                con.Open();
                string select = "insert into SGLFINAL(CodeRub,GLACCOUNT,GLDESCRIPTION,COSTCENTER,DESCRIPTION,SENS,VALEUR,IdSoc)"
               + " values('" + CodeRub + "','" + GLACCOUNT + "','" + GLDESCRIPTION + "','" + COSTCENTER + "','" + Desc + "','" + Sens + "'," + valeur.Replace(",", ".") + "," + IdSoc + ")";
                OleDbCommand cmd = new OleDbCommand(select, con);
                cmd.ExecuteNonQuery();
                etat = true;
            }

            catch (Exception ex)
            {
                etat = false;
            }
            finally
            {
                if (con != null)
                {
                    con.Close();
                    con = null;
                }
            }

            return etat;
        }

        public bool ModifierLeUser(int RefUser,string Login,string Pass)
        {
            OleDbConnection con = null;
            bool etat = false;
            try
            {
                con = new OleDbConnection(connection_string);
                con.Open();
                string select = "update  USERS set Login='" + Login + "',Pass='" + Pass + "'"
                + " where N°=" + RefUser + "";
                OleDbCommand cmd = new OleDbCommand(select, con);
                cmd.ExecuteNonQuery();
                etat = true;
            }

            catch (Exception)
            {
                etat = false;
            }
            finally
            {
                if (con != null)
                {
                    con.Close();
                    con = null;
                }
            }

            return etat;
        }

        public bool ModifierSociete(int Refer, string nom, string code)
        {
            OleDbConnection con = null;
            bool etat = false;
            try
            {
                con = new OleDbConnection(connection_string);
                con.Open();
                string select = "update  Societe set NomSoc='" + nom + "',CodeSoc='" + code + "'"
                + " where Num=" + Refer + "";
                OleDbCommand cmd = new OleDbCommand(select, con);
                cmd.ExecuteNonQuery();
                etat = true;
            }

            catch (Exception)
            {
                etat = false;
            }
            finally
            {
                if (con != null)
                {
                    con.Close();
                    con = null;
                }
            }

            return etat;
        }

        public bool ModifierRubric(int Refer, string code, string libelles)
        {
            OleDbConnection con = null;
            bool etat = false;
            try
            {
                con = new OleDbConnection(connection_string);
                con.Open();
                string select = "update  ListeDesRubrique set CodeRub='" + code + "',Libelles='" + libelles + "'"
                + " where Num=" + Refer + "";
                OleDbCommand cmd = new OleDbCommand(select, con);
                cmd.ExecuteNonQuery();
                etat = true;
            }

            catch (Exception)
            {
                etat = false;
            }
            finally
            {
                if (con != null)
                {
                    con.Close();
                    con = null;
                }
            }

            return etat;
        }

        public bool ModifierGlFile(int Refer, string CodeRub,string LibRub, string CodeCpt,string Libelles,string Sens)
        {
            OleDbConnection con = null;
            bool etat = false;
            try
            {
                con = new OleDbConnection(connection_string);
                con.Open();
                string select = "update  GLFILE set CodeRub='" + CodeRub + "',LibCodeRub='" + LibRub + "',CodeCpt='" + CodeCpt + "',Libelles='" + Libelles + "',Sens='" + Sens + "'"
                + " where Num=" + Refer + "";
                OleDbCommand cmd = new OleDbCommand(select, con);
                cmd.ExecuteNonQuery();
                etat = true;
            }

            catch (Exception)
            {
                etat = false;
            }
            finally
            {
                if (con != null)
                {
                    con.Close();
                    con = null;
                }
            }

            return etat;
        }

        public OleDbDataReader VerifierIdentif(string Login, string Pass)
        {
            OleDbConnection con = null;
            String ss = "select * from USERS where Login='"+Login+"'"
            +"and Pass='"+Pass+"'";
            con = con = new OleDbConnection(connection_string);
            con.Open();
            OleDbCommand cmdx = new OleDbCommand(ss, con);
            OleDbDataReader dr = cmdx.ExecuteReader();
            return dr;
        }

        public bool insertRaison(string rais, string codeSoc)
        {
            OleDbConnection con = null;
            bool etat = false;
            try
            {
                con = new OleDbConnection(connection_string);
                con.Open();
                string select = "insert into Societe (NomSoc,CodeSoc)"
               + " values('" + rais + "','" + codeSoc + "')";
                OleDbCommand cmd = new OleDbCommand(select, con);
                cmd.ExecuteNonQuery();
                etat = true;
            }

            catch (Exception ex)
            {
                etat = false;
            }
            finally
            {
                if (con != null)
                {
                    con.Close();
                    con = null;
                }
            }

            return etat;
        }

        public bool insertRubrique(string CodeLib, string Libelles,string idSoc,int NbMax)
        {
            OleDbConnection con = null;
            bool etat = false;
            try
            {
                con = new OleDbConnection(connection_string);
                con.Open();
                string select = "insert into ListeDesRubrique (CodeRub,Libelles,idSoc,OrdreSeq)"
               + " values('" + CodeLib + "','" + Libelles + "'," + idSoc + "," + NbMax + ")";
                OleDbCommand cmd = new OleDbCommand(select, con);
                cmd.ExecuteNonQuery();
                etat = true;
            }

            catch (Exception ex)
            {
                etat = false;
            }
            finally
            {
                if (con != null)
                {
                    con.Close();
                    con = null;
                }
            }

            return etat;
        }

        public String Conc(String type, int lentgh, String valeur)
        {
            int longueurManquante = lentgh - valeur.Length;
            if (longueurManquante > 0)
            {
                if (type == "x")
                {
                    for (int i = 0; i < longueurManquante; i++)
                    {
                        valeur = valeur + " ";
                    }
                }
                else
                {
                    for (int i = 0; i < longueurManquante; i++)
                    {
                        valeur = "0" + valeur;
                    }
                }

            }

            return valeur;

        }

        public String ConcChaine(String type, int lentgh, String valeur)
        {
            int longueurManquante = lentgh - valeur.Length;
            if (longueurManquante > 0)
            {
                if (type == "x")
                {
                    for (int i = 0; i < longueurManquante; i++)
                    {
                        valeur = " " +valeur;
                    }
                }

            }

            return valeur;

        }

        public bool modifIncrementation(string IdSoc, int NumIncrement)
        {
            OleDbConnection con = null;
            bool etat = false;
            try
            {
                con = new OleDbConnection(connection_string);
                con.Open();
                string select = "update  Increment set NumIncrement=" + NumIncrement + ""
                + " where IdSoc= '" + IdSoc + "'";
                OleDbCommand cmd = new OleDbCommand(select, con);
                cmd.ExecuteNonQuery();
                etat = true;
            }

            catch (Exception ex)
            {
                etat = false;
            }
            finally
            {
                if (con != null)
                {
                    con.Close();
                    con = null;
                }
            }
            return etat;
        }

        public bool insertIncrement(string codeSoc, int num)
        {
            OleDbConnection con = null;
            bool etat = false;
            try
            {
                con = new OleDbConnection(connection_string);
                con.Open();
                string select = "insert into Increment (IdSoc,NumIncrement)"
               + " values('" + codeSoc + "'," + num + ")";
                OleDbCommand cmd = new OleDbCommand(select, con);
                cmd.ExecuteNonQuery();
                etat = true;
            }

            catch (Exception ex)
            {
                etat = false;
            }
            finally
            {
                if (con != null)
                {
                    con.Close();
                    con = null;
                }
            }

            return etat;
        }

        public int recapIncrement(string num)
        {
            int nb = 0;
            OleDbConnection con = null;
            con = new OleDbConnection(connection_string);
            string select = "SELECT NumIncrement FROM Increment where CodeSoc = '" + num + "'";
            OleDbCommand da = new OleDbCommand(select, con);
            try
            {
                con.Open();
                nb = Convert.ToInt16(da.ExecuteScalar());

            }
            catch (Exception ex)
            {
                nb = 0;
            }
            finally
            {
                if (con != null)
                {
                    con.Close();
                    con = null;
                }
            }
            return nb;
        }

        public DataTable AffichageParam(string codeSoc)
        {
            DataTable dt = new DataTable();
            OleDbConnection con = null;
            try
            {
                con = new OleDbConnection(connection_string);
                OleDbDataAdapter adapt = new OleDbDataAdapter("select * from Parametre where CodeSoc='" + codeSoc + "'", con);
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

        public DataTable AffichageRUB(string codeSoc)
        {
            DataTable dt = new DataTable();
            OleDbConnection con = null;
            try
            {
                con = new OleDbConnection(connection_string);
                OleDbDataAdapter adapt = new OleDbDataAdapter("select * from ListeDesRubrique where CodeSoc='" + codeSoc + "'", con);
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

        public List<DataRaison> GetSociteList()
        {
            OleDbDataReader reader = null;
            OleDbCommand cmd = null;
            List<DataRaison> societes = new List<DataRaison>();
            string query = "SELECT * FROM Societe ORDER BY NomSoc ASC";
            try
            {
                cnx.Open();
                cmd = cnx.CreateCommand();
                cmd.Connection = cnx;
                cmd.CommandText = query;
                reader = cmd.ExecuteReader();
                DataRaison societe;
                while (reader.Read())
                {
                    societe = new DataRaison();
                    societe.Num = int.Parse(reader["Num"].ToString());
                    societe.Societe = reader["NomSoc"].ToString();
                    societe.CodeSoc = reader["CodeSoc"].ToString();
                    societes.Add(societe);
                }
            }
            catch (Exception ex)
            {
                //throw new Exception(ex.Message);
            }
            finally
            {
                if (reader != null)
                {
                    reader.Close();
                    reader = null;
                }
                if (cmd != null)
                {
                    cmd = null;
                }
                if (cnx != null)
                {
                    cnx.Close();
                }
            }
            return societes;
        }

        public DataRaison GetDetailsSociete(int IdSoc)
        {
            OleDbDataReader reader = null;
            OleDbCommand cmd = null;
            DataRaison societe = new DataRaison();

            string query = "SELECT * FROM Societe where Num = " + IdSoc;
            try
            {
                cnx.Open();
                cmd = cnx.CreateCommand();
                cmd.Connection = cnx;
                cmd.CommandText = query;
                reader = cmd.ExecuteReader();
                
                if (reader.Read())
                {
                  
                    societe.Num = int.Parse(reader["Num"].ToString());
                    societe.Societe = reader["NomSoc"].ToString();
                    societe.CodeSoc = reader["CodeSoc"].ToString();
                    societe.GrossToNetTableName = reader["GrossToNetTableName"].ToString();
                   
                }
            }
            catch (Exception ex)
            {
                //throw new Exception(ex.Message);
            }
            finally
            {
                if (reader != null)
                {
                    reader.Close();
                    reader = null;
                }
                if (cmd != null)
                {
                    cmd = null;
                }
                if (cnx != null)
                {
                    cnx.Close();
                }
            }
            return societe;
        }



        public bool insertDataFixe(string IDOracle, string IDLocale, string Civilite, string CiviliteAbrege, string CiviliteValeur, string Sexe, string SexeValeur,
           string Nom, string NomJeuneFille, string Prenom, string Adresse, string ComplAdresse, string Commune, string CodePostal,          
           string Tel, string CodePays, string IntitulePays, string Nationalite, string IntituleNationalite, string CIN,
           string DateExpirationSejour, string CodePaysNaiss, string DateNaiss, string ComuneNaiss, string SitFam, string SitFamVal,
           string NbreEnfant, string EnfRenseignes, string NumCNSS, string CleCNSS, string DateEmbauche, string DateDepart,
           string EtabSal, string TypeEntreEtab, string DateSortieEtab, string CodeConvColSal, string DeptSal, string IntituleDept,
           string SceSal, string UniteSal, string IntituleUnite, string CategSal, string CodeFonction, string EmploiOccupee,string NatContrat, string DateDebContrat,
           string BullModeleSal, string TypeSal, string SalBaseSal, string HorBaseSal, string CongeResteAnnPrec,
           string ModePaie, string CodeBque, string CodeGuichet, string NumCpt, string CleRib, string NomGuichet,
            string NomBque, string LibCpte, string MtAssVie, string EnfACharge, string ChefFamille, string AdhesionAsGr, string Champs1,
            string champs2, string Champs3, string Champs4, string Champs5, string Champs6, string RtCalcule, string CodeSce,
            string IntituleSce, string Date1,int RefSoc)
        {
            OleDbConnection con = null;
            bool etat = false;
            try
            {
                con = new OleDbConnection(connection_string);
                con.Open();
                string select = "insert into DATAFIXE (IDOracle,IDLocale,Civilite,CiviliteAbrege,CiviliteValeur,Sexe,SexeValeur,Nom,"
                + "NomJeuneFille,Prenom,Adresse,ComplAdresse,Commune,CodePostal,Tel,CodePays,IntitulePays,"
                + "Nationalite,IntituleNationalite,CIN,DateExpirationSejour,CodePaysNaiss,DateNaiss,ComuneNaiss,SitFam,SitFamVal,"
                + "NbreEnfant,EnfRenseignes,NumCNSS,CleCNSS,DateEmbauche,DateDepart,EtabSal,TypeEntreEtab,DateSortieEtab,CodeConvColSal,DeptSal,"
                + "IntituleDept,SceSal,UniteSal,IntituleUnite,CategSal,CodeFonction,EmploiOccupee,NatureContrat,DateDebContrat,BullModeleSal,TypeSal,SalBaseSal,"
                + "HorBaseSal,CongeResteAnnPrec,ModePaie,CodeBque,CodeGuichet,NumCpt,CleRib,NomGuichet,NomBque,LibCpte,MtAssVie,EnfACharge,ChefFamille,"
                + "AdhesionAsGr,Champs1,Champs2,Champs3,Champs4,Champs5,Champs6,RtCalcule,CodeSce,IntituleSce,Date1,RefSoc)"
               + " values('" + IDOracle + "','" + IDLocale + "','" + Civilite + "','" + CiviliteAbrege + "','" + CiviliteValeur + "','" + Sexe + "','" + SexeValeur + "','"+Nom+"',"
                + " '" + NomJeuneFille + "','" + Prenom + "','" + Adresse + "','" + ComplAdresse + "','" + Commune + "','" + CodePostal + "',"
                + " '" + Tel + "','" + CodePays + "','" + IntitulePays + "','" + Nationalite + "','" + IntituleNationalite + "','" + CIN + "',"
                + " '" + DateExpirationSejour + "','" + CodePaysNaiss + "','" + DateNaiss + "','" + ComuneNaiss + "','" + SitFam + "','" + SitFamVal + "',"
                + " '" + NbreEnfant + "','" + EnfRenseignes + "','" + NumCNSS + "','" + CleCNSS + "','" + DateEmbauche + "','" + DateDepart + "',"
                + " '" + EtabSal + "','" + TypeEntreEtab + "','" + DateSortieEtab + "','" + CodeConvColSal + "','" + DeptSal + "','" + IntituleDept + "',"
                + " '" + SceSal + "','" + UniteSal + "','" + IntituleUnite + "','" + CategSal + "','" + CodeFonction + "','" + EmploiOccupee + "','"+NatContrat+"','" + DateDebContrat + "','" + BullModeleSal + "',"
                + " '" + TypeSal + "','" + SalBaseSal + "','" + HorBaseSal + "','" + CongeResteAnnPrec + "','" + ModePaie + "','" + CodeBque + "',"
                + " '" + CodeGuichet + "','" + NumCpt + "','" + CleRib + "','" + NomGuichet + "','" + NomBque + "','" + LibCpte + "','"+MtAssVie+"','"+EnfACharge+"',"
                + " '" + ChefFamille + "','" + AdhesionAsGr + "','" + Champs1 + "','" + champs2 + "','" + Champs3 + "','" + Champs4 + "','" + Champs5 + "','" + Champs6 + "',"
                + " '" + RtCalcule + "','" + CodeSce + "','" + IntituleSce + "','"+Date1+"',"+RefSoc+")";
                OleDbCommand cmd = new OleDbCommand(select, con);
                cmd.ExecuteNonQuery();
                etat = true;
            }

            catch (Exception ex)
            {
                etat = false;
            }
            finally
            {
                if (con != null)
                {
                    con.Close();
                    con = null;
                }
            }

            return etat;
        }

        public string NomDuMois(int Mois)
        {
            string Selstatus = string.Empty;
            switch (Mois)
            {
                case 1:
                    Selstatus = "Janvier";
                    break;
                case 2:
                    Selstatus = "Fevrier";
                    break;
                case 3:
                    Selstatus = "Mars";
                    break;
                case 4:
                    Selstatus = "Avril";
                    break;
                case 5:
                    Selstatus = "Mai";
                    break;
                case 6:
                    Selstatus = "Juin";
                    break;
                case 7:
                    Selstatus = "Juillet";
                    break;
                case 8:
                    Selstatus = "Aout";
                    break;
                case 9:
                    Selstatus = "Septembre";
                    break;
                case 10:
                    Selstatus = "Octobre";
                    break;
                case 11:
                    Selstatus = "Novembre";
                    break;
                case 12:
                    Selstatus = "Decembre";
                    break;
                default:
                    Selstatus = "";
                    break;
            }

            return Selstatus;
        }

        public string NomDuMoisEnglish(int Mois)
        {
            string Selstatus = string.Empty;
            switch (Mois)
            {
                case 1:
                    Selstatus = "January";
                    break;
                case 2:
                    Selstatus = "February";
                    break;
                case 3:
                    Selstatus = "March";
                    break;
                case 4:
                    Selstatus = "April";
                    break;
                case 5:
                    Selstatus = "May";
                    break;
                case 6:
                    Selstatus = "June";
                    break;
                case 7:
                    Selstatus = "July";
                    break;
                case 8:
                    Selstatus = "August";
                    break;
                case 9:
                    Selstatus = "September";
                    break;
                case 10:
                    Selstatus = "October";
                    break;
                case 11:
                    Selstatus = "November";
                    break;
                case 12:
                    Selstatus = "December";
                    break;
                default:
                    Selstatus = "";
                    break;
            }

            return Selstatus;
        }


        public bool Delete_DATAFIXE()
        {
            System.Data.OleDb.OleDbConnection con;
            bool etat = false;
            con = new OleDbConnection(connection_string);
            con.Open();
            string Delete_Soc = @"DELETE FROM DATAFIXE";
            OleDbCommand Delete_T_Soc_Commande = new OleDbCommand(Delete_Soc, con);
            try
            {
                int count = Delete_T_Soc_Commande.ExecuteNonQuery();
                etat = true;
            }
            catch (OleDbException ex)
            {
                etat = false;
            }
            finally
            {
                con.Close();
            }

            return etat;

        }

        public DataTable RecupererData(int refSoc)
        {
            DataTable dt = new DataTable();
            OleDbConnection con = null;
            try
            {
                con = new OleDbConnection(connection_string);
                OleDbDataAdapter adapt = new OleDbDataAdapter("SELECT * from DATAFIXE F "
                +"where F.RefSoc="+refSoc+"", con);
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

        public DataTable RecupererListeRub(int refSoc)
        {
            DataTable dt = new DataTable();
            OleDbConnection con = null;
            try
            {
                con = new OleDbConnection(connection_string);
                OleDbDataAdapter adapt = new OleDbDataAdapter("select * from ListeDesRubrique where idSoc='"+refSoc+"'"
                + " ORDER BY OrdreSeq ASC", con);
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

        public DataTable ListeRubSVR(int refSoc,int cpt)
        {
            DataTable dt = new DataTable();
            OleDbConnection con = null;
            try
            {
                con = new OleDbConnection(connection_string);
                OleDbDataAdapter adapt = new OleDbDataAdapter("select * from ListeDesRubrique where idSoc='" + refSoc + "'"
                + " ORDER BY OrdreSeq ASC", con);
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

        public DataTable ListeSceSGL(string NomTab,int Mois,int Annee,string CodeSce,string IntituleSce)
        {
            DataTable dt = new DataTable();
            OleDbConnection con = null;
            try
            {
                con = new OleDbConnection(connection_string);
                string selection = "select distinct "+CodeSce+","+IntituleSce+" from " + NomTab + " where Mois="+Mois+" and Annee="+Annee+""
                + " ORDER BY "+CodeSce+" ASC";
                OleDbDataAdapter adapt = new OleDbDataAdapter(selection, con);
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

        public DataTable ListeOPR11(int Mois, int Annee)
        {
            DataTable dt = new DataTable();
            OleDbConnection con = null;
            try
            {
                con = new OleDbConnection(connection_string);
                string selection = "select * FROM GROSSTONETA016 where Mois=" + Mois + " and Annee=" + Annee + ""
                + " ORDER BY  Matricule ASC";
                OleDbDataAdapter adapt = new OleDbDataAdapter(selection, con);
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

        public DataTable ListeOPR8(int Mois, int Annee)
        {
            DataTable dt = new DataTable();
            OleDbConnection con = null;
            try
            {
                con = new OleDbConnection(connection_string);
                string selection = "select * FROM GROSSTONETA016 where Mois=" + Mois + " and Annee=" + Annee + ""
                + " ORDER BY  Matricule ASC";
                OleDbDataAdapter adapt = new OleDbDataAdapter(selection, con);
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

        public DataTable Rub11()
        {
            DataTable dt = new DataTable();
            OleDbConnection con = null;
            try
            {
                con = new OleDbConnection(connection_string);
                string selection = "select * FROM OPR11 ORDER BY NUM ASC";
                OleDbDataAdapter adapt = new OleDbDataAdapter(selection, con);
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

        public DataTable ListDeGlFileData(string COSTCENTER, string GLACCOUNT,string sens)
        {
            DataTable dt = new DataTable();
            OleDbConnection con = null;
            try
            {
                con = new OleDbConnection(connection_string);
                string select = "select GLACCOUNT,GLDESCRIPTION,"
                + "COSTCENTER,DESCRIPTION,SENS from SGLFINAL "
                + "where COSTCENTER='" + COSTCENTER + "' and GLACCOUNT='" + GLACCOUNT + "'"
                +" and SENS='"+sens+"'";
                OleDbDataAdapter adapt = new OleDbDataAdapter(select, con);
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

        public DataTable ListGlAccount()
        {
            DataTable dt = new DataTable();
            OleDbConnection con = null;
            try
            {
                con = new OleDbConnection(connection_string);
                OleDbDataAdapter adapt = new OleDbDataAdapter("select distinct COSTCENTER from SGLFINAL", con);
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

        public DataTable CodeCptF(string COSTCENTER)
        {
            DataTable dt = new DataTable();
            OleDbConnection con = null;
            try
            {
                con = new OleDbConnection(connection_string);
                OleDbDataAdapter adapt = new OleDbDataAdapter("select distinct GLACCOUNT,SENS from SGLFINAL where COSTCENTER='" + COSTCENTER + "'", con);
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

        public DataTable ListeRubCptSGL(int IdSoc)
        {
            DataTable dt = new DataTable();
            OleDbConnection con = null;
            try
            {
                con = new OleDbConnection(connection_string);
                string selection = "select distinct CodeRub,LibCodeRub,CodeCpt,Libelles,Sens from GLFILE where IdSoc=" + IdSoc + "";
                OleDbDataAdapter adapt = new OleDbDataAdapter(selection, con);
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

        public DataTable ListeEmployee(string table, int mois,int annee)
        {
            DataTable dt = new DataTable();
            OleDbConnection con = null;
            try
            {
                con = new OleDbConnection(connection_string);
                OleDbDataAdapter adapt = new OleDbDataAdapter("select * from " + table + " where Mois=" + mois + " and Annee=" + annee + "", con);
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

        public DataTable RecupererListe2330(int refSoc)
        {
            DataTable dt = new DataTable();
            OleDbConnection con = null;
            try
            {
                con = new OleDbConnection(connection_string);
                OleDbDataAdapter adapt = new OleDbDataAdapter("select * from ListeDesRubrique where idSoc='" + refSoc + "'"
                + " and CodeRub between '3100' and '3319' ORDER BY CodeRub ASC", con);
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

        public DataTable RecupererDATASGN(int refSoc,int MoisPe,int annee)
        {
            DataTable dt = new DataTable();
            OleDbConnection con = null;
            try
            {
                con = new OleDbConnection(connection_string);
                OleDbDataAdapter adapt = new OleDbDataAdapter("select * from GROSSTONETA001 where IdSoc=" + refSoc + " and Mois="+MoisPe+" and Annee="+annee+"", con);
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

        public DataTable RecupererDATASGNA002(int refSoc, int MoisPe, int annee)
        {
            DataTable dt = new DataTable();
            OleDbConnection con = null;
            try
            {
                con = new OleDbConnection(connection_string);
                OleDbDataAdapter adapt = new OleDbDataAdapter("select * from GROSSTONETA002 where IdSoc=" + refSoc + " and Mois=" + MoisPe + " and Annee=" + annee + "", con);
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

        public DataTable RecupererDATASGNA003(int refSoc, int MoisPe, int annee)
        {
            DataTable dt = new DataTable();
            OleDbConnection con = null;
            try
            {
                con = new OleDbConnection(connection_string);
                OleDbDataAdapter adapt = new OleDbDataAdapter("select * from GROSSTONETA003 where IdSoc=" + refSoc + " and Mois=" + MoisPe + " and Annee=" + annee + "", con);
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

        public DataTable RecupererDATASGNA004(int refSoc, int MoisPe, int annee)
        {
            DataTable dt = new DataTable();
            OleDbConnection con = null;
            try
            {
                con = new OleDbConnection(connection_string);
                OleDbDataAdapter adapt = new OleDbDataAdapter("select * from GROSSTONETA004 where IdSoc=" + refSoc + " and Mois=" + MoisPe + " and Annee=" + annee + "", con);
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


      


        public DataTable RecupererDATASGNA010(int refSoc, int MoisPe, int annee)
        {
            DataTable dt = new DataTable();
            OleDbConnection con = null;
            try
            {
                con = new OleDbConnection(connection_string);
                OleDbDataAdapter adapt = new OleDbDataAdapter("select * from GROSSTONETA010 where IdSoc=" + refSoc + " and Mois=" + MoisPe + " and Annee=" + annee + "", con);
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

        public DataTable RecupererDATASGNA013(int refSoc, int MoisPe, int annee)
        {
            DataTable dt = new DataTable();
            OleDbConnection con = null;
            try
            {
                con = new OleDbConnection(connection_string);
                OleDbDataAdapter adapt = new OleDbDataAdapter("select * from GROSSTONETA013 where IdSoc=" + refSoc + " and Mois=" + MoisPe + " and Annee=" + annee + "", con);
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

        public DataTable RecupererDATASGNA020(int refSoc, int MoisPe, int annee)
        {
            DataTable dt = new DataTable();
            OleDbConnection con = null;
            try
            {
                con = new OleDbConnection(connection_string);
                OleDbDataAdapter adapt = new OleDbDataAdapter("select * from GROSSTONETA020 where IdSoc=" + refSoc + " and Mois=" + MoisPe + " and Annee=" + annee + "", con);
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

        public DataTable RecupererDATASGNA016(int refSoc, int MoisPe, int annee)
        {
            DataTable dt = new DataTable();
            OleDbConnection con = null;
            try
            {
                con = new OleDbConnection(connection_string);
                OleDbDataAdapter adapt = new OleDbDataAdapter("select * from GROSSTONETA016 where IdSoc=" + refSoc + " and Mois=" + MoisPe + " and Annee=" + annee + "", con);
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

        public DataTable RecupererDATASGNA021(int refSoc, int MoisPe, int annee)
        {
            DataTable dt = new DataTable();
            OleDbConnection con = null;
            try
            {
                con = new OleDbConnection(connection_string);
                OleDbDataAdapter adapt = new OleDbDataAdapter("select * from GROSSTONETA021 where IdSoc=" + refSoc + " and Mois=" + MoisPe + " and Annee=" + annee + "", con);
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

        public DataTable RecupererDATASGNA011(int refSoc, int MoisPe, int annee)
        {
            DataTable dt = new DataTable();
            OleDbConnection con = null;
            try
            {
                con = new OleDbConnection(connection_string);
                OleDbDataAdapter adapt = new OleDbDataAdapter("select * from GROSSTONETA011 where IdSoc=" + refSoc + " and Mois=" + MoisPe + " and Annee=" + annee + "", con);
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

        public DataTable RecupererDATASGNA014(int refSoc, int MoisPe, int annee)
        {
            DataTable dt = new DataTable();
            OleDbConnection con = null;
            try
            {
                con = new OleDbConnection(connection_string);
                OleDbDataAdapter adapt = new OleDbDataAdapter("select * from GROSSTONETA014 where IdSoc=" + refSoc + " and Mois=" + MoisPe + " and Annee=" + annee + "", con);
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

        public DataTable RecupererDATASGNA015(int refSoc, int MoisPe, int annee)
        {
            DataTable dt = new DataTable();
            OleDbConnection con = null;
            try
            {
                con = new OleDbConnection(connection_string);
                OleDbDataAdapter adapt = new OleDbDataAdapter("select * from GROSSTONETA015 where IdSoc=" + refSoc + " and Mois=" + MoisPe + " and Annee=" + annee + "", con);
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

        public DataTable RecupererDATASEP(int refSoc, int MoisPe,int annee)
        {
            DataTable dt = new DataTable();
            OleDbConnection con = null;
            try
            {
                con = new OleDbConnection(connection_string);
                OleDbDataAdapter adapt = new OleDbDataAdapter("select * from PAYMENTREPORT where RefSoc=" + refSoc + " and MoisPaie=" + MoisPe + " and AnneePaie="+annee+"", con);
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

        public DataTable RecupererListe7600(int refSoc)
        {
            DataTable dt = new DataTable();
            OleDbConnection con = null;
            try
            {
                con = new OleDbConnection(connection_string);
                OleDbDataAdapter adapt = new OleDbDataAdapter("select * from ListeDesRubrique where idSoc='" + refSoc + "'"
                + " and CodeRub ='7600'", con);
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

        public DataTable RecupererListe4813(int refSoc)
        {
            DataTable dt = new DataTable();
            OleDbConnection con = null;
            try
            {
                con = new OleDbConnection(connection_string);
                OleDbDataAdapter adapt = new OleDbDataAdapter("select * from ListeDesRubrique where idSoc='" + refSoc + "'"
                + " and CodeRub between '4800' and '4815'", con);
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

        public DataTable RecupererEmployee(int RefSoc)
        {
            DataTable dt = new DataTable();
            OleDbConnection con = null;
            try
            {
                con = new OleDbConnection(connection_string);
                OleDbDataAdapter adapt = new OleDbDataAdapter("SELECT DISTINCT IDORACLE from DATAFIXE where RefSoc="+RefSoc+"", con);
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

        public OleDbDataReader RecupererEmp(string Tab, int Mois, int Annee,string NatMat)
        {
            OleDbConnection con = null;
            String ss = "SELECT DISTINCT "+NatMat+",Nom,Prenom from "+Tab+" where Mois=" + Mois + " and Annee="+Annee+" order by "+NatMat+" ASC";
            con = con = new OleDbConnection(connection_string);
            con.Open();
            OleDbCommand cmdx = new OleDbCommand(ss, con);
            OleDbDataReader dr = cmdx.ExecuteReader();
            return dr;
        }

        public double BrutGross(int refSoc, int mois1)
        {
            double nb = 0.0;
            OleDbConnection con = null;
            con = new OleDbConnection(connection_string);
            string select = "SELECT SUM(BRUT) from GROSSTONETA001 "
                + "where Mois="+mois1+"";
            OleDbCommand da = new OleDbCommand(select, con);
            try
            {
                con.Open();
                nb = Convert.ToDouble(da.ExecuteScalar());

            }
            catch (Exception)
            {
                nb = 0.0;
            }
            finally
            {
                if (con != null)
                {
                    con.Close();
                    con = null;
                }
            }
            return nb;
        }

        public double ValeurFinalSens(string GLACCOUNT,string COSTCENTER,string SENS)
        {
            double nb = 0.0;
            OleDbConnection con = null;
            con = new OleDbConnection(connection_string);
            string select = "SELECT SUM(VALEUR) from SGLFINAL "
                + "where GLACCOUNT='" + GLACCOUNT + "' and COSTCENTER='" + COSTCENTER + "' and Sens='"+SENS+"'";
            OleDbCommand da = new OleDbCommand(select, con);
            try
            {
                con.Open();
                nb = Convert.ToDouble(da.ExecuteScalar());

            }
            catch (Exception)
            {
                nb = 0.0;
            }
            finally
            {
                if (con != null)
                {
                    con.Close();
                    con = null;
                }
            }
            return nb;
        }

        public double RecupererValeur(int refSoc, int mois,int annee,string Code,string NomTab)
        {
            double nb = 0.0;
            OleDbConnection con = null;
            con = new OleDbConnection(connection_string);
            string select = "SELECT SUM(" + Code + ") from "+NomTab+" where Mois=" + mois + " and Annee=" + annee + "";
            OleDbCommand da = new OleDbCommand(select, con);
            try
            {
                con.Open();
                nb = Convert.ToDouble(da.ExecuteScalar());

            }
            catch (Exception)
            {
                nb = 0.0;
            }
            finally
            {
                if (con != null)
                {
                    con.Close();
                    con = null;
                }
            }
            return nb;
        }

        public double RecupererSumRubrique(int refSoc, int mois, int annee, string Code, string NomTab)
        {
            double nb = 0.0;
            OleDbConnection con = null;
            con = new OleDbConnection(connection_string);
            string select = "SELECT SUM(" + Code + ") from " + NomTab + " where Mois between 1 and " + mois + " and Annee=" + annee + "";
            OleDbCommand da = new OleDbCommand(select, con);
            try
            {
                con.Open();
                nb = Convert.ToDouble(da.ExecuteScalar());

            }
            catch (Exception)
            {
                nb = 0.0;
            }
            finally
            {
                if (con != null)
                {
                    con.Close();
                    con = null;
                }
            }
            return nb;
        }

        public double RecupererVF(int refSoc, int mois, int annee, string Code, string NomTab, string Mat)
        {
             string NomMat=string.Empty;
            if ((refSoc == 4) || (refSoc == 3)||(refSoc==11)||(refSoc==9))
            {
                NomMat = "IdOracle";
            }
            else
            {
                NomMat = "Matricule";
            }
            double nb = 0.0;
            OleDbConnection con = null;
            con = new OleDbConnection(connection_string);
            string select = "SELECT SUM(" + Code + ") from " + NomTab + " where Mois=" + mois + " and Annee=" + annee + " and "+NomMat+"='" + Mat + "'";
            OleDbCommand da = new OleDbCommand(select, con);
            try
            {
                con.Open();
                nb = Convert.ToDouble(da.ExecuteScalar());

            }
            catch (Exception)
            {
                nb = 0.0;
            }
            finally
            {
                if (con != null)
                {
                    con.Close();
                    con = null;
                }
            }
            return nb;
        }

        public double RecupererSVE(int refSoc, int mois, int annee, string Code, string NomTab,string mat)
        {
            double nb = 0.0;
            OleDbConnection con = null;
            con = new OleDbConnection(connection_string);
            string select = "SELECT SUM(" + Code + ") from " + NomTab + " where IdSoc=" + refSoc + " and Mois=" + mois + " and Annee=" + annee + " and Matricule='"+mat+"'";
            OleDbCommand da = new OleDbCommand(select, con);
            try
            {
                con.Open();
                nb = Convert.ToDouble(da.ExecuteScalar());

            }
            catch (Exception)
            {
                nb = 0.0;
            }
            finally
            {
                if (con != null)
                {
                    con.Close();
                    con = null;
                }
            }
            return nb;
        }

        public double RecupererSVEIdOra(int refSoc, int mois, int annee, string Code, string NomTab, string mat)
        {
            double nb = 0.0;
            OleDbConnection con = null;
            con = new OleDbConnection(connection_string);
            string select = "SELECT SUM(" + Code + ") from " + NomTab + " where IdSoc=" + refSoc + " and Mois=" + mois + " and Annee=" + annee + " and IdOracle='" + mat + "'";
            OleDbCommand da = new OleDbCommand(select, con);
            try
            {
                con.Open();
                nb = Convert.ToDouble(da.ExecuteScalar());

            }
            catch (Exception)
            {
                nb = 0.0;
            }
            finally
            {
                if (con != null)
                {
                    con.Close();
                    con = null;
                }
            }
            return nb;
        }

        public double Gross1100(int refSoc, int mois1)
        {
            double nb = 0.0;
            OleDbConnection con = null;
            con = new OleDbConnection(connection_string);
            string select = "SELECT SUM(1100) from GROSSTONETA001 "
                + "where Mois=" + mois1 + "";
            OleDbCommand da = new OleDbCommand(select, con);
            try
            {
                con.Open();
                nb = Convert.ToDouble(da.ExecuteScalar());

            }
            catch (Exception)
            {
                nb = 0.0;
            }
            finally
            {
                if (con != null)
                {
                    con.Close();
                    con = null;
                }
            }
            return nb;
        }

        public double DataCodeMemo(string codeM,string Mat)
        {
            double nb = 0.0;
            OleDbConnection con = null;
            string NomMat = string.Empty;
            con = new OleDbConnection(connection_string);
            string select = "SELECT VALEUR from CODEMEMO where Matricule='" + Mat + "' and CODEMEMO='" + codeM + "'";
            OleDbCommand da = new OleDbCommand(select, con);
            try
            {
                con.Open();
                nb = Convert.ToDouble(da.ExecuteScalar());

            }
            catch (Exception)
            {
                nb = 0.0;
            }
            finally
            {
                if (con != null)
                {
                    con.Close();
                    con = null;
                }
            }
            return nb;
        }

        public double SUMCodeMemo(string codeM)
        {
            double nb = 0.0;
            OleDbConnection con = null;
            con = new OleDbConnection(connection_string);
            string select = "SELECT SUM(VALEUR) from CODEMEMO where CODEMEMO='" + codeM + "'";
            OleDbCommand da = new OleDbCommand(select, con);
            try
            {
                con.Open();
                nb = Convert.ToDouble(da.ExecuteScalar());

            }
            catch (Exception)
            {
                nb = 0.0;
            }
            finally
            {
                if (con != null)
                {
                    con.Close();
                    con = null;
                }
            }
            return nb;
        }

        public double CnssEmplyeurF(string codeM, string NomTab, int Mois, int Annee,string Defalquetion, string CodeSce)
        {
            double nb = 0.0;
            OleDbConnection con = null;
            con = new OleDbConnection(connection_string);
            string select = "SELECT SUM(" + codeM + ") from " + NomTab + " where Mois=" + Mois + " and Annee=" + Annee + " and " + Defalquetion + "='" + CodeSce + "'";
            OleDbCommand da = new OleDbCommand(select, con);
            try
            {
                con.Open();
                nb = Convert.ToDouble(da.ExecuteScalar());

            }
            catch (Exception)
            {
                nb = 0.0;
            }
            finally
            {
                if (con != null)
                {
                    con.Close();
                    con = null;
                }
            }
            return nb;
        }

        public double SUMRetenue(string codeM,string NomTab,int Mois,int Annee,string Mat)
        {
            double nb = 0.0;
            OleDbConnection con = null;
            con = new OleDbConnection(connection_string);
            string select = "SELECT "+codeM+" from "+NomTab+" where Mois="+Mois+" and Annee="+Annee+" and Matricule='"+Mat+"'";
            OleDbCommand da = new OleDbCommand(select, con);
            try
            {
                con.Open();
                nb = Convert.ToDouble(da.ExecuteScalar());

            }
            catch (Exception)
            {
                nb = 0.0;
            }
            finally
            {
                if (con != null)
                {
                    con.Close();
                    con = null;
                }
            }
            return nb;
        }

        public int NbSal()
        {
            int nb = 0;
            OleDbConnection con = null;
            con = new OleDbConnection(connection_string);
            string select = "SELECT COUNT(*) from DATAFIXE";
            OleDbCommand da = new OleDbCommand(select, con);
            try
            {
                con.Open();
                nb = Convert.ToInt16(da.ExecuteScalar());

            }
            catch (Exception)
            {
                nb = 0;
            }
            finally
            {
                if (con != null)
                {
                    con.Close();
                    con = null;
                }
            }
            return nb;
        }

        public string RIBBGMOISPREC(string Mat,int MoisPr,int AnneePr)
        {
            string nb = string.Empty;
            OleDbConnection con = null;
            con = new OleDbConnection(connection_string);
            string select = "SELECT RIB FROM GROSSTONETA016 where Matricule = '" + Mat + "' and Mois="+MoisPr+" and Annee="+AnneePr+"";
            OleDbCommand da = new OleDbCommand(select, con);
            try
            {
                con.Open();
                nb = Convert.ToString(da.ExecuteScalar());

            }
            catch (Exception ex)
            {
                nb = string.Empty;
            }
            finally
            {
                if (con != null)
                {
                    con.Close();
                    con = null;
                }
            }
            return nb;
        }

        public string ValRubA016(string Mat, int MoisPr, int AnneePr,string Rub)
        {
            string nb = string.Empty;
            OleDbConnection con = null;
            con = new OleDbConnection(connection_string);
            string select = "SELECT "+Rub+" FROM GROSSTONETA016 where Matricule = '" + Mat + "' and Mois=" + MoisPr + " and Annee=" + AnneePr + "";
            OleDbCommand da = new OleDbCommand(select, con);
            try
            {
                con.Open();
                nb = Convert.ToString(da.ExecuteScalar());

            }
            catch (Exception ex)
            {
                nb = string.Empty;
            }
            finally
            {
                if (con != null)
                {
                    con.Close();
                    con = null;
                }
            }
            return nb;
        }

        public string NETPAIEBGMOISPREC(string Mat, int MoisPr, int AnneePr)
        {
            string nb = string.Empty;
            OleDbConnection con = null;
            con = new OleDbConnection(connection_string);
            string select = "SELECT NETPAIE FROM GROSSTONETA016 where Matricule = '" + Mat + "' and Mois=" + MoisPr + " and Annee=" + AnneePr + "";
            OleDbCommand da = new OleDbCommand(select, con);
            try
            {
                con.Open();
                nb = Convert.ToString(da.ExecuteScalar());

            }
            catch (Exception ex)
            {
                nb = string.Empty;
            }
            finally
            {
                if (con != null)
                {
                    con.Close();
                    con = null;
                }
            }
            return nb;
        }

        public string SALBASEBGMOISPREC(string Mat, int MoisPr, int AnneePr)
        {
            string nb = string.Empty;
            OleDbConnection con = null;
            con = new OleDbConnection(connection_string);
            string select = "SELECT A1000 FROM GROSSTONETA016 where Matricule = '" + Mat + "' and Mois=" + MoisPr + " and Annee=" + AnneePr + "";
            OleDbCommand da = new OleDbCommand(select, con);
            try
            {
                con.Open();
                nb = Convert.ToString(da.ExecuteScalar());

            }
            catch (Exception ex)
            {
                nb = string.Empty;
            }
            finally
            {
                if (con != null)
                {
                    con.Close();
                    con = null;
                }
            }
            return nb;
        }

        public string NatureContrat(string CodeSal)
        {
            string nb = string.Empty;
            OleDbConnection con = null;
            con = new OleDbConnection(connection_string);
            string select = "SELECT NatureContrat FROM GROSSTONET where MatSal = '" + CodeSal + "'";
            OleDbCommand da = new OleDbCommand(select, con);
            try
            {
                con.Open();
                nb = Convert.ToString(da.ExecuteScalar());

            }
            catch (Exception ex)
            {
                nb = string.Empty;
            }
            finally
            {
                if (con != null)
                {
                    con.Close();
                    con = null;
                }
            }
            return nb;
        }

        public double TesterSensCont(string GLACCOUNT,string Sens,string Sce)
        {
            double nb = 0;
            OleDbConnection con = null;
            con = new OleDbConnection(connection_string);
            string select = "SELECT SUM(VALEUR) AS VAL FROM SGLFINAL WHERE GLACCOUNT='" + GLACCOUNT + "' AND SENS='"+Sens+"' and COSTCENTER='"+Sce+"'";
            OleDbCommand da = new OleDbCommand(select, con);
            try
            {
                con.Open();
                nb = Convert.ToDouble(da.ExecuteScalar());

            }
            catch (Exception ex)
            {
                nb = 0;
            }
            finally
            {
                if (con != null)
                {
                    con.Close();
                    con = null;
                }
            }
            return nb;
        }

        public string SalNet(string CodeSal,string NomTab,int mois,int annee)
        {
            string nb = string.Empty;
            string NatMat = string.Empty;
            OleDbConnection con = null;
            if ((NomTab == "GROSSTONETA020") || (NomTab == "GROSSTONETA014"))
            {
                NatMat = "IdOracle";
            }
            else
            {
                NatMat = "Matricule";
            }
            con = new OleDbConnection(connection_string);
            string select = "SELECT NETPAIE FROM " + NomTab + " where "+NatMat+" = '" + CodeSal + "' and Mois="+mois+" and Annee="+annee+"";
            OleDbCommand da = new OleDbCommand(select, con);
            try
            {
                con.Open();
                nb = Convert.ToString(da.ExecuteScalar());

            }
            catch (Exception ex)
            {
                nb = string.Empty;
            }
            finally
            {
                if (con != null)
                {
                    con.Close();
                    con = null;
                }
            }
            return nb;
        }

        public double ValeurCodeMemo(string Valeur,string Table,string Mat,int Mois,int Annee)
        {
            double nb = 0.0;
            OleDbConnection con = null;
            con = new OleDbConnection(connection_string);
            string select = "SELECT "+Valeur+" FROM " + Table + " where Matricule = '" + Mat + "' and Mois=" + Mois + " and Annee=" + Annee + "";
            OleDbCommand da = new OleDbCommand(select, con);
            try
            {
                con.Open();
                nb = Convert.ToDouble(da.ExecuteScalar());

            }
            catch (Exception ex)
            {
                nb = 0.0;
            }
            finally
            {
                if (con != null)
                {
                    con.Close();
                    con = null;
                }
            }
            return nb;
        }

        public double ValeurCodeGLFILE(string Valeur, string Table,int Mois, int Annee,string CodeSce,string CentreC)
        {
            double nb = 0.0;
            OleDbConnection con = null;
            con = new OleDbConnection(connection_string);
            string select = "SELECT SUM(" + Valeur + ") FROM " + Table + " where Mois=" + Mois + " and Annee=" + Annee + " and "+CentreC+"='"+CodeSce+"'";
            OleDbCommand da = new OleDbCommand(select, con);
            try
            {
                con.Open();
                nb = Convert.ToDouble(da.ExecuteScalar());

            }
            catch (Exception ex)
            {
                nb = 0.0;
            }
            finally
            {
                if (con != null)
                {
                    con.Close();
                    con = null;
                }
            }
            return nb;
        }

        public string Avance(string CodeSal, string NomTab, int mois, int annee)
        {
            string nb = string.Empty;
            OleDbConnection con = null;
            con = new OleDbConnection(connection_string);
            string select = "SELECT A6410 FROM " + NomTab + " where Matricule = '" + CodeSal + "' and Mois=" + mois + " and Annee=" + annee + "";
            OleDbCommand da = new OleDbCommand(select, con);
            try
            {
                con.Open();
                nb = Convert.ToString(da.ExecuteScalar());

            }
            catch (Exception ex)
            {
                nb = string.Empty;
            }
            finally
            {
                if (con != null)
                {
                    con.Close();
                    con = null;
                }
            }
            return nb;
        }

        public int OrdreSeq(int IdSoc)
        {
            int nb = 0;
            OleDbConnection con = null;
            con = new OleDbConnection(connection_string);
            string select = "SELECT MAX(OrdreSeq) FROM ListeDesRubrique where idSoc='"+IdSoc+"'";
            OleDbCommand da = new OleDbCommand(select, con);
            try
            {
                con.Open();
                nb = Convert.ToInt16(da.ExecuteScalar());

            }
            catch (Exception ex)
            {
                nb = 0;
            }
            finally
            {
                if (con != null)
                {
                    con.Close();
                    con = null;
                }
            }
            return nb;
        }

        public bool insertGrossToNetA001(string matricule, string nom, string prenom, string DateEmbauche, string DateDept, string CodeSce, string IntituleSce,
                               string NatureContrat, string Emploi, string A1100, string CL01, string A1200, string A2200, string A2100, string A2311,
                                string A2320, string A3318, string A3110, string A4311, string A4330, string A3111, string A3210, string A3317, string A3319, string A4100,
                                string A4171,string A4130, string A4223, string A4380, string A4711, string A4717, string A5100, string A7600, string A4811, string A4812, string A4813,
                                string BRUT, string AjusCot, string A9112, string A6110, string A6120, string A6210, string A6230, string A9121, string A6310, string A6311,
                                string A6614, string A6410, string A6520,string A6510, string A6530, string A6890, string RUB_DEDSA, string NETPAIE, string A61109, string A61209, string A6150,
                                string A62109, string A62309, string A6320, string A6330,string A8000, string RUB_COTEPY,string A6490, string MoisPe, string NumSoc, int Mois, int Annee,
                                string A4712,string A4226)
        {
            OleDbConnection con = null;
            bool etat = false;
            try
            {
                con = new OleDbConnection(connection_string);
                con.Open();
                string select = "insert into GROSSTONETA001 (Matricule,Nom,Prenom,DateEmbauche,DateDepart,CodeSce,IntituleSce,NatureContrat,Emploi,A1100,CL01,"
                + "A1200,A2200,A2100,A2311,A2320,A3318,A3110,A4311,A4330,A3111,A3210,A3317,A3319,A4100,A4171,A4130,A4223,A4380,A4711,A4717,A5100,A7600,A4811,A4812,A4813,BRUT,AJUSCOT,A9112,A6110,"
                + "A6120,A6210,A6230,A9121,A6310,A6311,A6614,A6410,A6520,A6510,A6530,A6890,A8000,RUB_DEDSA,NETPAIE,A61109,A61209,A6150,A62109,A62309,A6320,A6330,RUB_COTEPY,A6490,MOIS_PAIE,IdSoc,"
                + "Mois,Annee,A4712,A4226)"
               + " values('" + matricule + "','" + nom + "','" + prenom + "','" + DateEmbauche + "','" + DateDept + "','" + CodeSce + "','" + IntituleSce + "'"
                + ",'" + NatureContrat + "','" + Emploi + "','" + A1100 + "','" + CL01 + "','" + A1200 + "','" + A2200 + "' ,'" + A2100 + "'"
                + ",'" + A2311 + "','" + A2320 + "','" + A3318 + "','" + A3110 + "','" + A4311 + "','" + A4330 + "','" + A3111 + "'"
                + ",'" + A3210 + "','" + A3317 + "','" + A3319 + "','" + A4100 + "','" + A4171 + "','" + A4130 + "','" + A4223 + "','" + A4380 + "'"
                + ",'" + A4711 + "','" + A4717 + "','" + A5100 + "','" + A7600 + "','" + A4811 + "','" + A4812 + "','" + A4813 + "'"
                + ",'" + BRUT + "','" + AjusCot + "','" + A9112 + "','" + A6110 + "','" + A6120 + "','" + A6210 + "','" + A6230 + "'"
                + ",'" + A9121 + "','" + A6310 + "','" + A6311 + "','" + A6614 + "','" + A6410 + "','" + A6520 + "','" + A6510 + "'"
                + ",'" + A6530 + "','" + A6890 + "','"+A8000+"','" + RUB_DEDSA + "','" + NETPAIE + "','" + A61109 + "','" + A61209 + "','" + A6150 + "'"
                + ",'" + A62109 + "','" + A62309 + "','" + A6320 + "','" + A6330 + "','" + RUB_COTEPY + "','" + A6490 + "','" + MoisPe + "'," + NumSoc + "," + Mois + "," + Annee + ""
                + ",'"+A4712 + "','"+A4226+"')";
                OleDbCommand cmd = new OleDbCommand(select, con);
                cmd.ExecuteNonQuery();
                etat = true;
            }

            catch (Exception ex)
            {
                etat = false;
            }
            finally
            {
                if (con != null)
                {
                    con.Close();
                    con = null;
                }
            }

            return etat;
        }

        public bool insertGrossToNetA002(string Matricule, string Nom, string Prenom, string DateEmbauche, string DateDepart, string CodeSce, string IntituleSce
           , string NatureContrat, string Emploi, string SalBase, string NbJrsPres, string NbHrPres, string SalBaseJrsTrav, string IndemTrans, string IndemPres,
           string CongePayer, string CongSpecio, string CongAbcNonRem, string PrimExcepFix, string CongAPayer, string HSupp125, string HSupp150, string HSupp200,
           string PrimAstreint, string PrimDim, string PrimDiff, string PrimDiffFor, string PrimParr, string PrimFormation, string PrimException, string PrimObjectif, string PrimNaiss, string PrimMariage,
           string PrimDeces, string PrimAid, string RappelEltRec, string RappelEltNonRec, string AvantNatTR, string RappelAvNatTr, string AvanNatAssGrp, string IndemPreavis,
           string GratifFinS, string IndemLicen,string A4815, string SalBrut, string SalBrutCot, string CNSSEmp, string RtnAssGrp, string SalImposable, string IRPP, string Redev,
           string ContConj, string AvancSal, string PretSoc, string RtnTiketRes, string AutreDed, string DedAvanNat, string DedAvanNatTR, string DedAvanNatAssGr,string A8800,
           string TotDedEmp, string SalNetPay,string CNSSEmply,string AcciTrav,string AssGrEmp,string totContPat, string MOIS_PAIE, string IdSoc, int Mois, int annee,string A5100,
            string A7116,string A6850)
        {
            OleDbConnection con = null;
            bool etat = false;
            try
            {
                con = new OleDbConnection(connection_string);
                con.Open();
                string select = "insert into GROSSTONETA002 (Matricule,Nom,Prenom,DateEmbauche,DateDepart,CodeSce,IntituleSce,NatureContrat,Emploi,A1100,PRESC1J,"
                + "PRESC2H,A1000,A2200,A2100,A4170,A4173,A4174,A3710,A4171,A4110,A4111,A4113,A4130,A4140,A4217,A4218,A4220,A4221,A4222,A4340,A4711,A4712,A4715,A4717,CL25,CL24,"
                + "A7113,A7114,A7115,A4811,A4812,A4813,A4815,BRUT,A9112,A6110,A6210,A9121,A6310,A6311,A6614,A6410,A6520,A6430,A6490,A6830,A6831,A6840,A8800,RUB_DEDSA,NETPAIE,"
                + "A61109,A6150,A71159,RUB_COTEPY,MOIS_PAIE,IdSoc,Mois,Annee,A5100,A7116,A6850)"
               + " values('" + Matricule + "','" + Nom + "','" + Prenom + "','" + DateEmbauche + "','" + DateDepart + "','" + CodeSce + "','" + IntituleSce + "'"
                + ",'" + NatureContrat + "','" + Emploi + "','" + SalBase + "','" + NbJrsPres + "','" + NbHrPres + "','" + SalBaseJrsTrav + "' ,'" + IndemTrans + "'"
                + ",'" + IndemPres + "','" + CongePayer + "','" + CongSpecio + "','" + CongAbcNonRem + "','" + PrimExcepFix + "','" + CongAPayer + "','" + HSupp125 + "'"
                + ",'" + HSupp150 + "','" + HSupp200 + "','" + PrimAstreint + "','" + PrimDim + "','"+PrimDiff+"','" + PrimDiffFor + "','" + PrimParr + "','" + PrimFormation + "','" + PrimException + "'"
                + ",'" + PrimObjectif + "','" + PrimNaiss + "','" + PrimMariage + "','" + PrimDeces + "','" + PrimAid + "','" + RappelEltRec + "','" + RappelEltNonRec + "'"
                + ",'" + AvantNatTR + "','" + RappelAvNatTr + "','" + AvanNatAssGrp + "','" + IndemPreavis + "','" + GratifFinS + "','" + IndemLicen + "','"+A4815+"','" + SalBrut + "'"
                + ",'" + SalBrutCot + "','" + CNSSEmp + "','" + RtnAssGrp + "','" + SalImposable + "','" + IRPP + "','" + Redev + "','" + ContConj + "'"
                + ",'" + AvancSal + "','" + PretSoc + "','" + RtnTiketRes + "','" + AutreDed + "','" + DedAvanNat + "','" + DedAvanNatTR + "','" + DedAvanNatAssGr + "','" + A8800 + "'"
                + ",'" + TotDedEmp + "','" + SalNetPay + "','" + CNSSEmply + "','" + AcciTrav + "','" + AssGrEmp + "','" + totContPat + "','" + MOIS_PAIE + "'," + IdSoc + "," + Mois + "," + annee + ",'"+A5100+"'"
                + ",'"+A7116+"','"+A6850+"')";
                OleDbCommand cmd = new OleDbCommand(select, con);
                cmd.ExecuteNonQuery();
                etat = true;
            }

            catch (Exception ex)
            {
                etat = false;
            }
            finally
            {
                if (con != null)
                {
                    con.Close();
                    con = null;
                }
            }

            return etat;
        }

        public bool insertGrossToNetA003(string Matricule, string Nom, string Prenom, string DateEmbauche, string DateDepart, string CodeSce, string IntituleSce
          , string NatureContrat, string Emploi, string SalBase, string CL01, string A1200, string A2200, string A2100, string A2360,
          string A2330, string A3412, string A3317, string A3712, string A3710, string A3711, string A3713, string A4112,
          string A4113, string A4120, string A4222,string A4216, string A4361, string A4215, string A4320, string A4370, string A4430, string A7113, string A4171,
          string BRUT, string AjusCot, string A9112, string A6110, string A6120, string A9310, string A6310, string A6311,
          string A6340, string A6210, string A6410, string A6430, string A6520, string A6510, string A6530, string A6830,string A6490, string RUB_DEDSA,
          string NETPAIE, string A6110A, string A6120A, string A6150, string A6210A, string RUB_COTEPY,string MOIS_PAIE, string IdSoc, int Mois, int annee,
            string A4812,string A5100,string A4380 )
        {
            OleDbConnection con = null;
            bool etat = false;
            try
            {
                con = new OleDbConnection(connection_string);
                con.Open();
                string select = "insert into GROSSTONETA003 (Matricule,Nom,Prenom,DateEmbauche,DateDepart,CodeSce,IntituleSce,NatureContrat,Emploi,A1100,CL01,"
                + "A1200,A2200,A2100,A2360,A2330,A3412,A3317,A3712,A3710,A3711,A3713,A4112,A4113,A4120,A4222,A4216,A4361,A4215,A4320,A4370,A4430,A7113,A4171,BRUT,AjusCot,"
                + "A9112,A6110,A6120,A9310,A6310,A6311,A6340,A6210,A6410,A6430,A6520,A6510,A6530,A6830,A6490,RUB_DEDSA,NETPAIE,A61109,A61209,A6150,A62109,"
                + " RUB_COTEPY,MoisPe,IdSoc,Mois,Annee,A4812,A5100,A4380)"
               + " values('" + Matricule + "','" + Nom + "','" + Prenom + "','" + DateEmbauche + "','" + DateDepart + "','" + CodeSce + "','" + IntituleSce + "'"
                + ",'" + NatureContrat + "','" + Emploi + "','" + SalBase + "','" + CL01 + "','" + A1200 + "','" + A2200 + "' ,'" + A2100 + "'"
                + ",'" + A2360 + "','" + A2330 + "','" + A3412 + "','" + A3317 + "','" + A3712 + "','" + A3710 + "','" + A3711 + "'"
                + ",'" + A3713 + "','" + A4112 + "','" + A4113 + "','" + A4120 + "','" + A4222 + "','" + A4216 + "','" + A4361 + "','" + A4215 + "','" + A4320 + "','" + A4370 + "'"
                + ",'" + A4430 + "','" + A7113 + "','" + A4171 + "','" + BRUT + "','" + AjusCot + "','" + A9112 + "','" + A6110 + "'"
                + ",'" + A6120 + "','" + A9310 + "','" + A6310 + "','" + A6311 + "','" + A6340 + "','" + A6210 + "','" + A6410 + "'"
                + ",'" + A6430 + "','" + A6520 + "','" + A6510 + "','" + A6530 + "','" + A6830 + "','"+A6490+"','" + RUB_DEDSA + "','" + NETPAIE + "'"
                + ",'" + A6110A + "','" + A6120A + "','" + A6150 + "','" + A6210A + "','" + RUB_COTEPY + "','" + MOIS_PAIE + "'," + IdSoc + "," + Mois + "," + annee + ",'"+A4812+"','"+A5100+"','"+ A4380 +"')";
                OleDbCommand cmd = new OleDbCommand(select, con);
                cmd.ExecuteNonQuery();
                etat = true;
            }

            catch (Exception ex)
            {
                etat = false;
            }
            finally
            {
                if (con != null)
                {
                    con.Close();
                    con = null;
                }
            }

            return etat;
        }

        public bool insertGrossToNetA004(string Matricule, string Nom, string Prenom, string DateEmbauche, string DateDepart, string CodeSce, string IntituleSce
         , string NatureContrat, string Emploi, string SalBase, string CL01, string A1100, string A2200, string A2100, string A3411,string A3410,
         string A3412, string A3317, string A3712, string A3710,string A3711,string A3713,string A4112,string A4113,string A4120,string A4216,
         string A4222,string A4361,string A4215,string A4320,string A4370,string A4430,string A7113,string BRUT,string AjusCot,
         string A9112,string A6110,string A6210,string A9310,string A6310,string A6311,string A6340,string A6210A,string A6410,
         string A6430,string A6520,string A6510,string A6530,string A6830,string A6490,string RUB_DEDSA,string NETPAIE,string A6110A,
            string A6120,string A6150,string A6210AA,string RUB_COTEPY,string MOIS_PAIE, string IdSoc, int Mois, int annee,string A4811,string A4171,string A4380,string A4812)
        {
            OleDbConnection con = null;
            bool etat = false;
            try
            {
                con = new OleDbConnection(connection_string);
                con.Open();
                string select = "insert into GROSSTONETA004(Matricule,Nom,Prenom,DateEmbauche,DateDepart,CodeSce,IntituleSce,NatureContrat,Emploi,A1100,CL01,"
                + "A1200,A2200,A2100,A3411,A3410,A3412,A3317,A3712,A3710,A3711,A3713,A4112,A4113,A4120,A4216,A4222,A4361,A4215,A4320,A4370,A4430,A7113,BRUT,AjusCot,A9112,"
                + "A6110,A6210,A9310,A6310,A6311,A6340,A62109,A6410,A6430,A6520,A6510,A6530,A6830,A6490,RUB_DEDSA,NETPAIE,A61109,A6120,A6150,A621099,"
                + "RUB_COTEPY,MoisPe,IdSoc,Mois,Annee,A4811,A4171,A4380,A4812)"
               + " values('" + Matricule + "','" + Nom + "','" + Prenom + "','" + DateEmbauche + "','" + DateDepart + "','" + CodeSce + "','" + IntituleSce + "'"
                + ",'" + NatureContrat + "','" + Emploi + "','" + SalBase + "','" + CL01 + "','" + A1100 + "','" + A2200 + "' ,'" + A2100 + "'"
                + ",'" + A3411 + "','" + A3410 + "','" + A3412 + "','" + A3317 + "','" + A3712 + "','" + A3710 + "','" + A3711 + "'"
                + ",'" + A3713 + "','" + A4112 + "','" + A4113 + "','" + A4120 + "','" + A4216 + "','" + A4222 + "','" + A4361 + "','" + A4215 + "','" + A4320 + "','" + A4370 + "'"
                + ",'" + A4430 + "','" + A7113 + "','" + BRUT + "','" + AjusCot + "','" + A9112 + "','" + A6110 + "'"
                + ",'" + A6210 + "','" + A9310 + "','" + A6310 + "','" + A6311 + "','" + A6340 + "','" + A6210A + "','" + A6410 + "'"
                + ",'" + A6430 + "','" + A6520 + "','" + A6510 + "','" + A6530 + "','" + A6830 + "','" + A6490 + "','" + RUB_DEDSA + "','" + NETPAIE + "'"
                + ",'" + A6110A + "','" + A6120 + "','" + A6150 + "','" + A6210AA + "','" + RUB_COTEPY + "','" + MOIS_PAIE + "'," + IdSoc + "," + Mois + ","
                + "" + annee + ",'"+A4811+"','"+A4171+"','"+A4380+"','"+ A4812 +"')";
                OleDbCommand cmd = new OleDbCommand(select, con);
                cmd.ExecuteNonQuery();
                etat = true;
            }

            catch (Exception ex)
            {
                etat = false;
            }
            finally
            {
                if (con != null)
                {
                    con.Close();
                    con = null;
                }
            }

            return etat;
        }

        public bool insertGrossToNetA021(
            string matricule,
            string nom,
            string prenom,
                string Datedembauchesociete ,
                        string Datededepartsociete ,
                        string Intituledepartement ,
                        string CodeCostcenter ,
                        string Intitulecostcenter ,
                        string Intituledelanatureducontrat ,
                        string Emploioccupe ,
                        string Salairedebase ,
                        string Nbjoursdepresence ,
                        string SalairedebaseJtrvaille ,
                        string Inddepresence ,
                        string InddeTransport ,
                        string PrimeComplementairedePresence ,
                        string IndemnitedeVoiture ,
                        string IndFixederepresentation ,
                        string Indemnitedexpatriation ,
                      string  HS125HS01,
                      string  ValeurHS01,
                      string  HS150HS02,
                       string ValeurHS02,
                       string HS175HS06	,
                      string  ValeurHS06,
                      string  HS200HS03,
                       string ValeurHS03,
                       string HsupplementairenuitHS04,
                      string  valeurHsuplementairenuitHS04,
                      string  primeLVP,
                      string  PrimeSIP,
                      string  PrimeAOS,
                       string Bonus,
                     string   Compensationsupport,
                      string  Rappelsursalaire,
                      string  Rappelprimedetransport,
                      string  PrimeperequationRC,
                      string  PrimeperequationAV,
                       string PrimeAid	,
                      string  Primemariage,
                      string  Primedenaissance,
                      string  Congespayes,
                      string  Indemnitedepreavis,
                      string  Gratificationfindeservice,
                      string  Indemnitedelicenciement,
                      string  Avticketsrestaurant,
                      string  AvAssurancemaladie,
                      string  Avassuranceretraitecomplementaire,
                      string  Avcarburant,
                      string  Avvoiture,
                      string  Avscolariteenfants,
                     string   Avutiliteexpat,
                     string   Avlogement,
                       string SalaireBrut,
                       string BrutEricsson,
                      string  CNSSEmploye,
                      string  AssuranceretraitecomplemCNRPSemploye,
                      string  AssuranceretraiteComplementaire,
                     string   Salaireimposable,
                      string  IRPP,
                      string  Redevance1,
                      string  PrêtCNSS,
                      string  AutresRetenuesursalaire,
                      string  DeductionAvTR,
                      string  DeductionAvassurancemaladie	,
                      string  DeductionAvassurancecomplementaire	,
                      string  DeductionAvVoiture,
                      string  DeductionAvCarburant,
                      string  DeductionAvloyer,
                       string DeductionAvIndExpat	,
                       string Deductionutiliteexpat,
                      string  NetàPayer,
                      string  CNSSEmployeur,
                      string  Accidentdetravail,
                      string  AssurancemaladieEmployeur,
                      string  AssuranceComplementaireEmployeur,
                      string  AssuranceretraitcompleCNRPSemplye,
                      string  TFP,
                      string  FOPROLOS,
                       string TotalContributionEmployeur,
                       string Moisdepaie 



        
        )
        {
            OleDbConnection con = null;
            bool etat = false;
            try
            {
                con = new OleDbConnection(connection_string);
                con.Open();
                int IdSoc = 14;
                int Mois = 12;
                int Annee = 2015;

                string select = "insert into GROSSTONETA021 ";

                select = select + "(matricule"
                      + ",nom"
                       + " , prenom"
               + ", DateEmbauche"
               + ", DateDepart"
              + ", Intituledepartement"
              + ", CodeCostcenter"
              + ", Intitulecostcenter"
              + ", Intituledelanatureducontrat"
              + ", Emploioccupe"
              + ", Salairedebase"
              + ", Nbjoursdepresence"
              + ", SalairedebaseJtrvaille"
              + ", Inddepresence"
              + ", InddeTransport"
              + ", PrimeComplementairedePresence"
              + ", IndemnitedeVoiture"
              + ", IndFixederepresentation"
              + ", Indemnitedexpatriation"
            + ", HS125HS01"
            + ", ValeurHS01"
            + ", HS150HS02"
             + ", ValeurHS02"
             + ", HS175HS06"
            + ", ValeurHS06"
            + ", HS200HS03"
             + ", ValeurHS03"
             + ", HsupplementairenuitHS04"
            + ", valeurHsuplementairenuitHS04"
            + ", primeLVP"
            + ", PrimeSIP"
            + ", PrimeAOS"
             + ", Bonus"
           + ", Compensationsupport"
            + ", Rappelsursalaire"
            + ", Rappelprimedetransport"
            + ", PrimeperequationRC"
            + ", PrimeperequationAV"
             + ", PrimeAid"
            + ", Primemariage"
            + ", Primedenaissance"
            + ", Congespayes"
            + ", Indemnitedepreavis"
            + ", Gratificationfindeservice"
            + ", Indemnitedelicenciement"
            + ", Avticketsrestaurant"
            + ", AvAssurancemaladie"
            + ", Avassuranceretraitecomplementaire"
            + ", Avcarburant"
            + ", Avvoiture"
            + ", Avscolariteenfants"
           + ", Avutiliteexpat"
           + ", Avlogement"
             + ", SalaireBrut"
             + ", BrutEricsson"
            + ", CNSSEmploye"
            + ", AssuranceretraitecomplemCNRPSemploye"
            + ", AssuranceretraiteComplementaire"
           + ", Salaireimposable"
            + ", IRPP"
            + ", Redevance1"
            + ", PrêtCNSS"
            + ", AutresRetenuesursalaire"
            + ", DeductionAvTR"
            + ", DeductionAvassurancemaladie"
            + ", DeductionAvassurancecomplementaire"
            + ", DeductionAvVoiture"
            + ", DeductionAvCarburant"
            + ", DeductionAvloyer"
             + ", DeductionAvIndExpat"
             + ", Deductionutiliteexpat"
            + ", NetàPayer"
            + ", CNSSEmployeur"
            + ", Accidentdetravail"
            + ", AssurancemaladieEmployeur"
            + ", AssuranceComplementaireEmployeur"
            + ", AssuranceretraitcompleCNRPSemplye"
            + ", TFP"
            + ", FOPROLOS"
             + ", TotalContributionEmployeur"
             + ", Moisdepaie,IdSoc,Mois,Annee)";
            


                select = select + "VALUES ('" + matricule
                      + "','" + nom
                        + "','" + prenom
               + "','" + Datedembauchesociete
               + "','" + Datededepartsociete
               + "','" + Intituledepartement
               + "','" + CodeCostcenter
               + "','" + Intitulecostcenter
               + "','" + Intituledelanatureducontrat
               + "','" + Emploioccupe
               + "','" + Salairedebase
               + "','" + Nbjoursdepresence
               + "','" + SalairedebaseJtrvaille
               + "','" + Inddepresence
               + "','" + InddeTransport
               + "','" + PrimeComplementairedePresence
               + "','" + IndemnitedeVoiture
               + "','" + IndFixederepresentation
               + "','" + Indemnitedexpatriation
             + "','" + HS125HS01
             + "','" + ValeurHS01
             + "','" + HS150HS02
              + "','" + ValeurHS02
              + "','" + HS175HS06
             + "','" + ValeurHS06
             + "','" + HS200HS03
              + "','" + ValeurHS03
              + "','" + HsupplementairenuitHS04
             + "','" + valeurHsuplementairenuitHS04
             + "','" + primeLVP
             + "','" + PrimeSIP
             + "','" + PrimeAOS
              + "','" + Bonus
            + "','" + Compensationsupport
             + "','" + Rappelsursalaire
             + "','" + Rappelprimedetransport
             + "','" + PrimeperequationRC
             + "','" + PrimeperequationAV
              + "','" + PrimeAid
             + "','" + Primemariage
             + "','" + Primedenaissance
             + "','" + Congespayes
             + "','" + Indemnitedepreavis
             + "','" + Gratificationfindeservice
             + "','" + Indemnitedelicenciement
             + "','" + Avticketsrestaurant
             + "','" + AvAssurancemaladie
             + "','" + Avassuranceretraitecomplementaire
             + "','" + Avcarburant
             + "','" + Avvoiture
             + "','" + Avscolariteenfants
            + "','" + Avutiliteexpat
            + "','" + Avlogement
              + "','" + SalaireBrut
              + "','" + BrutEricsson
             + "','" + CNSSEmploye
             + "','" + AssuranceretraitecomplemCNRPSemploye
             + "','" + AssuranceretraiteComplementaire
            + "','" + Salaireimposable
             + "','" + IRPP
             + "','" + Redevance1
             + "','" + PrêtCNSS
             + "','" + AutresRetenuesursalaire
             + "','" + DeductionAvTR
             + "','" + DeductionAvassurancemaladie
             + "','" + DeductionAvassurancecomplementaire
             + "','" + DeductionAvVoiture
             + "','" + DeductionAvCarburant
             + "','" + DeductionAvloyer
              + "','" + DeductionAvIndExpat
              + "','" + Deductionutiliteexpat
             + "','" + NetàPayer
             + "','" + CNSSEmployeur
             + "','" + Accidentdetravail
             + "','" + AssurancemaladieEmployeur
             + "','" + AssuranceComplementaireEmployeur
             + "','" + AssuranceretraitcompleCNRPSemplye
             + "','" + TFP
             + "','" + FOPROLOS
              + "','" + TotalContributionEmployeur
              + "','" + Moisdepaie
              + "',"+ IdSoc.ToString() + "," + Mois.ToString() + "," + Annee.ToString() + ")";


                OleDbCommand cmd = new OleDbCommand(select, con);
                cmd.ExecuteNonQuery();
                etat = true;
            }

            catch (Exception ex)
            {
                etat = false;
            }
            finally
            {
                if (con != null)
                {
                    con.Close();
                    con = null;
                }
            }

            return etat;
        }
        public bool insertGrossToNetA011(string Matricule, string Nom, string Prenom, string DateEmbauche, string DateDepart, string IntituleDept,string IntituleSce,
           string IntituleCat,string NatContrat,string Emploi,string SalBaseSal,string CL01,string A1200,string A2100,string A2200,string A2311,string A3111,
            string A3413,string A3410,string A4351,string A4220,string A4211,string A4369,string A4360,string A4380,string A4717,string A4716,string A4110,string A4111,
            string A4113,string A4114,string A4812,string A4171,string BRUT,string A6110,string A9121,string A6310,string A6311,string TotDecEmp,string NETPAIE,string A6110A,
            string A6150,string A6320,string A6330,string RUB_EPYC,string TotCostLab,string MOIS_PAIE, string IdSoc, int Mois, int annee,string CodeDept)
        {
            OleDbConnection con = null;
            bool etat = false;
            try
            {
                con = new OleDbConnection(connection_string);
                con.Open();
                string select = "insert into GROSSTONETA011(Matricule,Nom,Prenom,DateEmbauche,DateDepart,IntituleDept,IntituleSce,IntituleCat,NatContrat,Emploi,A1100,"
                + "CL01,A1200,A2100,A2200,A2311,A3111,A3413,A3410,A4351,A4220,A4211,A4369,A4360,A4380,A4717,A4716,A4110,A4111,A4113,A4114,A4812,A4171,BRUT,A6110,A9121,A6310,A6311,RUB_DEDSA,"
                + "NETPAIE,A61109,A6150,A6320,A6330,RUB_COTEPY,TotCostLab,MoisPe,IdSoc,Mois,Annee,CodeDept)"
               + " values('" + Matricule + "','" + Nom + "','" + Prenom + "','" + DateEmbauche + "','" + DateDepart + "','" + IntituleDept + "','" + IntituleSce + "'"
                + ",'" + IntituleCat + "','" + NatContrat + "','" + Emploi + "','" + SalBaseSal + "','" + CL01 + "','" + A1200 + "' ,'" + A2100 + "'"
                + ",'" + A2200 + "','" + A2311 + "','" + A3111 + "','" + A3413 + "','"+A3410+"','" + A4351 + "','" + A4220 + "'"
                + ",'" + A4211 + "','" + A4369 + "','" + A4360 + "','" + A4380 + "','" + A4717 + "','" + A4716 + "','" + A4110 + "','" + A4111 + "','" + A4113 + "'"
                + ",'" + A4114 + "','"+A4812+"','"+A4171+"','" + BRUT + "','" + A6110 + "','" + A9121 + "','" + A6310 + "','" + A6311 + "'"
                + ",'" + TotDecEmp + "','" + NETPAIE + "','" + A6110A + "','" + A6150 + "','" + A6320 + "','" + A6330 + "','" + RUB_EPYC + "'"
                + ",'" + TotCostLab + "','" + MOIS_PAIE + "'," + IdSoc + "," + Mois + "," + annee + ",'"+CodeDept+"')";
                OleDbCommand cmd = new OleDbCommand(select, con);
                cmd.ExecuteNonQuery();
                etat = true;
            }

            catch (Exception ex)
            {
                etat = false;
            }
            finally
            {
                if (con != null)
                {
                    con.Close();
                    con = null;
                }
            }

            return etat;
        }

        public bool insertGrossToNetA015(string Matricule, string Nom, string Prenom, string DateEmbauche, string DateDepart, string CodeDept,
            string IntituleDept, string NatContrat, string Emploi, string SalBaseAnn,string SALBASE_M,string CL01,string A1200,string A3410,
            string A3110,string A3711,string A4222,string A4362,string A4311,string A4380,string A4711,string A4717,string A4718,string A4364,string A4171,
            string A5100,string A7111,string A7112,string A7118,string BRUT,string A6110,string A6120,string A9121,string A6310,string A6311,
            string A8000,string A6850,string A6220,string RUB_DEDSA,string NETPAIE,string A6110A,string A6120A,string A6150,string RUB_EPYC,
            string MOIS_PAIE, string IdSoc, int Mois, int annee)
        {
            OleDbConnection con = null;
            bool etat = false;
            try
            {
                con = new OleDbConnection(connection_string);
                con.Open();
                string select = "insert into GROSSTONETA015(Matricule,Nom,Prenom,DateEmbauche,DateDepart,CodeDept,IntituleDept,NatContrat,Emploi,SalBaseAnn,A1100"
                + ",CL01,A1200,A3410,A3110,A3711,A4222,A4362,A4311,A4380,A4711,A4717,A4718,A4364,A4171,A5100,A7111,A7112,A7118,BRUT,A6110,A6120,A9121,A6310,A6311,A8000,A6850"
                + ",A6220,RUB_DEDSA,NETPAIE,A61109,A61209,A6150,RUB_COTEPY,MoisPe,IdSoc,Mois,Annee)"
               + " values('" + Matricule + "','" + Nom + "','" + Prenom + "','" + DateEmbauche + "','" + DateDepart + "','" + CodeDept + "','" + IntituleDept + "'"
                + ",'" + NatContrat + "','" + Emploi + "','" + SalBaseAnn + "','" + SALBASE_M + "','" + CL01 + "','" + A1200 + "' ,'" + A3410 + "'"
                + ",'" + A3110 + "','" + A3711 + "','" + A4222 + "','" + A4362 + "','" + A4311 + "','" + A4380 + "'"
                + ",'" + A4711 + "','" + A4717 + "','" + A4718 + "','" + A4364 + "','" + A4171 + "','" + A5100 + "','" + A7111 + "','" + A7112 + "','" + A7118 + "','" + BRUT + "'"
                + ",'" + A6110 + "','" + A6120 + "','" + A9121 + "','" + A6310 + "','" + A6311 + "','"+A8000+"','" + A6850 + "'"
                + ",'" + A6220 + "','" + RUB_DEDSA + "','" + NETPAIE + "','" + A6110A + "','" + A6120A + "','" + A6150 + "','" + RUB_EPYC + "'"
                + ",'" + MOIS_PAIE + "'," + IdSoc + "," + Mois + "," + annee + ")";
                OleDbCommand cmd = new OleDbCommand(select, con);
                cmd.ExecuteNonQuery();
                etat = true;
            }

            catch (Exception ex)
            {
                etat = false;
            }
            finally
            {
                if (con != null)
                {
                    con.Close();
                    con = null;
                }
            }

            return etat;
        }


        public bool insertPaymentReport(string Matricule, string Nom, string Prenom, string DateEmbauche, string DateDepart, string ModePe, string InfoBank, string NetPaie,
            int RefSoc, int MoisPaie, int AnneePaie)
        {
            OleDbConnection con = null;
            bool etat = false;
            try
            {
                con = new OleDbConnection(connection_string);
                con.Open();
                string select = "insert into PAYMENTREPORT (Matricule,Nom,Prenom,DateEmbauche,DateDepart,ModePe,InfoBank,NetPaie,RefSoc,MoisPaie,AnneePaie)"
               + " values('" + Matricule + "','" + Nom + "','" + Prenom + "','" + DateEmbauche + "','" + DateDepart + "','" + ModePe + "','" + InfoBank + "'"
                + ",'" + NetPaie + "'," + RefSoc + "," + MoisPaie + "," + AnneePaie + ")";
                OleDbCommand cmd = new OleDbCommand(select, con);
                cmd.ExecuteNonQuery();
                etat = true;
            }

            catch (Exception ex)
            {
                etat = false;
            }
            finally
            {
                if (con != null)
                {
                    con.Close();
                    con = null;
                }
            }

            return etat;
        }

        public bool insertPaymentReport(string CODEMEMO, string VALEUR, string Matricule,string idSoc)
        {
            OleDbConnection con = null;
            bool etat = false;
            try
            {
                con = new OleDbConnection(connection_string);
                con.Open();
                string select = "insert into CODEMEMO (CODEMEMO,VALEUR,Matricule,IdSoc)"
               + " values('" + CODEMEMO + "'," + VALEUR.Replace(",",".") + ",'" + Matricule + "','" + idSoc + "')";
                OleDbCommand cmd = new OleDbCommand(select, con);
                cmd.ExecuteNonQuery();
                etat = true;
            }

            catch (Exception ex)
            {
                etat = false;
            }
            finally
            {
                if (con != null)
                {
                    con.Close();
                    con = null;
                }
            }

            return etat;
        }

        public int NumberSalary(int mois,int annee,string NomTab)
        {
            int nb = 0;
            OleDbConnection con = null;
            con = new OleDbConnection(connection_string);
            string select = "SELECT Count(*) FROM " + NomTab + " where Mois = " + mois + " and Annee="+annee+"";
            OleDbCommand da = new OleDbCommand(select, con);
            try
            {
                con.Open();
                nb = Convert.ToInt16(da.ExecuteScalar());

            }
            catch (Exception ex)
            {
                nb = 0;
            }
            finally
            {
                if (con != null)
                {
                    con.Close();
                    con = null;
                }
            }
            return nb;
        }

        public double SumNetPay(int mois, int annee, string NomTab)
        {
            double nb = 0;
            OleDbConnection con = null;
            con = new OleDbConnection(connection_string);
            string select = "SELECT SUM(NETPAIE) FROM " + NomTab + " where Mois = " + mois + " and Annee=" + annee + "";
            OleDbCommand da = new OleDbCommand(select, con);
            try
            {
                con.Open();
                nb = Convert.ToDouble(da.ExecuteScalar());

            }
            catch (Exception ex)
            {
                nb = 0;
            }
            finally
            {
                if (con != null)
                {
                    con.Close();
                    con = null;
                }
            }
            return nb;
        }

        public double SumBrutPay(int mois, int annee, string NomTab)
        {
            double nb = 0;
            OleDbConnection con = null;
            con = new OleDbConnection(connection_string);
            string select = "SELECT SUM(BRUT) FROM " + NomTab + " where Mois = " + mois + " and Annee=" + annee + "";
            OleDbCommand da = new OleDbCommand(select, con);
            try
            {
                con.Open();
                nb = Convert.ToDouble(da.ExecuteScalar());

            }
            catch (Exception ex)
            {
                nb = 0;
            }
            finally
            {
                if (con != null)
                {
                    con.Close();
                    con = null;
                }
            }
            return nb;
        }

        public double SumNetPayA021(int mois, int annee, string NomTab)
        {
            double nb = 0;
            OleDbConnection con = null;
            con = new OleDbConnection(connection_string);
            string select = "SELECT SUM(NetàPayer) FROM " + NomTab + " where Mois = " + mois + " and Annee=" + annee + "";
            OleDbCommand da = new OleDbCommand(select, con);
            try
            {
                con.Open();
                nb = Convert.ToDouble(da.ExecuteScalar());

            }
            catch (Exception ex)
            {
                nb = 0;
            }
            finally
            {
                if (con != null)
                {
                    con.Close();
                    con = null;
                }
            }
            return nb;
        }

        public double SumBrutPayA021(int mois, int annee, string NomTab)
        {
            double nb = 0;
            OleDbConnection con = null;
            con = new OleDbConnection(connection_string);
            string select = "SELECT SUM(Salairedebase) FROM " + NomTab + " where Mois = " + mois + " and Annee=" + annee + "";
            OleDbCommand da = new OleDbCommand(select, con);
            try
            {
                con.Open();
                nb = Convert.ToDouble(da.ExecuteScalar());

            }
            catch (Exception ex)
            {
                nb = 0;
            }
            finally
            {
                if (con != null)
                {
                    con.Close();
                    con = null;
                }
            }
            return nb;
        }




        public bool insertGrossToNetA010(string IdOracle,string Matricule, string Nom, string Prenom, string DateEmbauche, string DateDepart,string Emploi,
           string NatureContrat,string CodeDept,string Dept,string HorBaseSal,string BullMod,string Categorie,string A1100,string A1300,string A2100,
            string A2200,string A3310,string ABS_CONGNR,string AbsNonRem,string RUB_ARR,string A4373,string HS01,string A4110,string HS02,string A4111,
            string HS03,string A4113,string HS06,string A4112,string HS04,string A4120,string HS07,string A4372,string HS05,string A4114,string HA04,
            string HA06,string RUB_JCHOM,string HA07,string RUB_TRP,string HA09,string A5500,string A4222,string A4371,string A4130,string A4135,
            string A2361,string A4816,string A4810,string HA01,string A4811,string A4812,string A5100,string A5400,string A5900,string A5200,
            string A6510,string A3309,string A4817,string A4818,string A4313,string A4312,string TESTSTC,string A4800,string PRES8,string PRES2,
            string A7115,string BRUT,string A6110,string A9310,string A6310,string A6311,string A6490,string A6491,string A6492,string A6840,
            string A6610,string A6611,string A6613,string A6612,string A6410,string A6511,string A6614,string RUB_DEDSA,string NETPAIE,string A6110A,
            string A6150,string A7115A,string RUB_COTEPY,string ModPe,string MoisPe,int IdSoc,int Mois,int Annee,string A4175,string A4374)
        {
            OleDbConnection con = null;
            bool etat = false;
            try
            {
                con = new OleDbConnection(connection_string);
                con.Open();
                string select = "insert into GROSSTONETA010(IdOracle,Matricule,Nom,Prenom,DateEmbauche,DateDepart,Emploi,NatureContrat,CodeDept,IntituleDept,HorBaseSal,"
                    + "BullMod,Categorie,A1100,A1300,A2100,A2200,A3310,ABS_CONGNR,A4174,RUB_ARR,A4373,HS01,A4110,HS02,A4111,HS03,A4113,HS06,A4112,HS04,A4120,"
                + "HS07,A4372,HS05,A4114,HA04,HA06,RUB_JCHOM,HA07,A4176,HA09,A5500,A4222,A4371,A4130,A4135,A2361,A4816,A4810,HA01,A4811,A4812,A5100,A5400,A5900,"
                + "A5200,A6510,A3309,A4817,A4818,A4313,A4312,TESTSTC,A4800,PRES8,PRES2,A7115,BRUT,A6110,A9310,A6310,A6311,A6490,A6491,A6492,A6840,A6610,A6611,"
                + "A6613,A6612,A6410,A6511,A6614,RUB_DEDSA,NETPAIE,A61109,A6150,A71159,RUB_COTEPY,ModPe,MoisPe,IdSoc,Mois,Annee,A4175,A4374)"
               + " values('" + IdOracle + "','" + Matricule + "','" + Nom + "','" + Prenom + "','" + DateEmbauche + "','" + DateDepart + "','" + Emploi + "'"
                + ",'" + NatureContrat + "','" + CodeDept + "','" + Dept + "','" + HorBaseSal + "','" + BullMod + "','" + Categorie + "' ,'" + A1100 + "'"
                + ",'" + A1300 + "','" + A2100 + "','" + A2200 + "','" + A3310 + "','" + ABS_CONGNR + "','" + AbsNonRem + "','" + RUB_ARR + "'"
                + ",'" + A4373 + "','" + HS01 + "','" + A4110 + "','" + HS02 + "','" + A4111 + "','" + HS03 + "','" + A4113 + "','" + HS06 + "','" + A4112 + "'"
                + ",'" + HS04 + "','" + A4120 + "','" + HS07 + "','" + A4372 + "','" + HS05 + "','" + A4114 + "'"
                + ",'" + HA04 + "','" + HA06 + "','" + RUB_JCHOM + "','" + HA07 + "','" + RUB_TRP + "','" + HA09 + "','" + A5500 + "'"
                + ",'" + A4222 + "','" + A4371 + "','" + A4130 + "','" + A4135 + "','" + A2361 + "','" + A4816 + "','" + A4810 + "'"
                + ",'" + HA01 + "','" + A4811 + "','" + A4812 + "','" + A5100 + "','" + A5400 + "','" + A5900 + "','" + A5200 + "','" + A6510 + "','" + A3309 + "'"
                + ",'" + A4817 + "','" + A4818 + "','" + A4313 + "','" + A4312 + "','" + TESTSTC + "','" + A4800 + "','" + PRES8 + "','" + PRES2 + "'"
                + ",'" + A7115 + "','" + BRUT + "','" + A6110 + "','" + A9310 + "','" + A6310 + "','" + A6311 + "','" + A6490 + "','" + A6491 + "','" + A6492 + "'"
                + ",'" + A6840 + "','" + A6610 + "','" + A6611 + "','" + A6613 + "','" + A6612 + "','" + A6410 + "','" + A6511 + "','" + A6614 + "','" + RUB_DEDSA + "'"
                + ",'" + NETPAIE + "','" + A6110A + "','" + A6150 + "','" + A7115A + "','" + RUB_COTEPY + "','" + ModPe + "','" + MoisPe + "'," + IdSoc + "," + Mois + "," + Annee + ",'"+A4175+"','"+A4374+"')";
                OleDbCommand cmd = new OleDbCommand(select, con);
                cmd.ExecuteNonQuery();
                etat = true;
            }

            catch (Exception ex)
            {
                etat = false;
            }
            finally
            {
                if (con != null)
                {
                    con.Close();
                    con = null;
                }
            }

            return etat;
        }

        public bool insertGrossToNetA013(string IdOracle,string Matricule,string Nom,string Prenom, string DateEmbauche,
                       string DateDept,string Emploi,string NatureContrat,string CodeDept,string Dept, string HorBaseSal,string BullMod,
                       string Categorie,string ModPe,string A1100,string A1300,string A1310,string A2100,string A2200,string A3310,string A3110,
                       string ABS_CONGNR,string A4174,string HS01,string A4110,string HS02,string A4111,string HS03,string A4113,string HS04,string A4120,
                       string HS07,string A4372,string HS05,string A4114,string HA04,string HA06,string A4175,string HA07,string A4176,string A4130,string A4135,
                       string A2361,string A4222,string A3309,string A4373,string A4371,string A4817,string A4818,string A4313,
                       string A4311,string HA09,string A5500,string A5100,string A5400,string A5200,string A5900,string HA01,string A4816,string A4810,string A4811,
                       string A4812,string TESTSTC,string A4800, string PRES8,string A7113,string CL23,string A7116,string A7115,
                       string BRUT,string A6110,string A9310,string A6310,string A6311,string A6614,string A6410,string A6490,string A6491,string A6492,string A6511,
                       string A6430,string A6431,string A6840,string RUB_DEDSA,string NETPAIE,string A61109,string A6150,string A71159,
                       string RUB_COTEPY,string MoisPe,int NumSoc,int Mois,int Annee,string A4374)
        {
            OleDbConnection con = null;
            bool etat = false;
            try
            {
                con = new OleDbConnection(connection_string);
                con.Open();
                string select = "insert into GROSSTONETA013(IdOracle, Matricule, Nom, Prenom, DateEmbauche,"
                       + "DateDepart, Emploi, NatureContrat, CodeDept, IntituleDept, HorBaseSal, BullMod,"
                       + "Categorie, ModPaie, A1100, A1300, A1310, A2100, A2200, A3310, A3110, ABS_CONGNR, A4174, HS01, A4110, HS02, A4111, HS03, A4113, HS04, A4120,"
                       +"HS07, A4372, HS05, A4114, HA04, HA06, A4175, HA07, A4176, A4130, A4135, A2361, A4222, A3309, A4373, A4371, A4817, A4818, A4313,"
                       +"A4311, HA09, A5500, A5100, A5400, A5200, A5900, HA01, A4816, A4810, A4811, A4812, TESTSTC, A4800, PRES8, A7113, CL23, A7116, A7115,"
                       +" BRUT, A6110, A9310, A6310, A6311, A6614, A6410, A6490, A6491, A6492, A6511, A6430, A6431, A6840, RUB_DEDSA, NETPAIE, A61109, A6150, A71159,"
                       + " RUB_COTEPY, MOIS_PAIE, IdSoc, Mois, Annee,A4374)"
               + " values('" + IdOracle + "','" + Matricule + "','" + Nom + "','" + Prenom + "','" + DateEmbauche + "','" + DateDept + "','" + Emploi + "'"
                + ",'" + NatureContrat + "','" + CodeDept + "','" + Dept + "','" + HorBaseSal + "','" + BullMod + "','" + Categorie + "' ,'"+ModPe+"','" + A1100 + "'"
                + ",'" + A1300 + "','" + A1310 + "','" + A2100 + "','" + A2200 + "','" + A3310 + "','"+A3110+"','" + ABS_CONGNR + "','" + A4174 + "'"
                + ",'" + HS01 + "','" + A4110 + "','" + HS02 + "','" + A4111 + "','" + HS03 + "','" + A4113 + "','" + HS04 + "','" + A4120 + "','" + HS07 + "'"
                + ",'" + A4372 + "','" + HS05 + "','" + A4114 + "','" + HA04 + "','" + HA06 + "','" + A4175 + "'"
                + ",'" + HA07 + "','"+A4176+"','"+A4130+"','" + A4135 + "','" + A2361 + "','" + A4222 + "','" + A3309 + "','" + A4373 + "'"
                + ",'" + A4371 + "','" + A4817 + "','" + A4818 + "','" + A4313 + "','"+A4311+"','" + HA09 + "','" + A5500 + "','" + A5100 + "'"
                + ",'" + A5400 + "','" + A5200 + "','" + A5900 + "','" + HA01 + "','" + A4816 + "','" + A4810 + "','" + A4811 + "','" + A4812 + "','" + TESTSTC + "'"
                + ",'" + A4800 + "','" + PRES8 + "','" + A7113 + "','" + CL23 + "','" + A7116 + "','" + A7115 + "','"+BRUT+"','" + A6110 + "','" + A9310 + "'"
                + ",'" + A6310 + "','" + A6311 + "','" + A6614 + "','"+A6410+"','" + A6490 + "','" + A6491 + "','" + A6492 + "','" + A6511 + "','"+A6430+"','" + A6431 + "','"+A6840+"'"
                + ",'" + RUB_DEDSA + "','" + NETPAIE + "','" + A61109 + "','" + A6150 + "','" + A71159 + "','" + RUB_COTEPY + "','" + MoisPe + "'"
                + "," + NumSoc + "," + Mois + "," + Annee + ",'"+A4374+"')";
                OleDbCommand cmd = new OleDbCommand(select, con);
                cmd.ExecuteNonQuery();
                etat = true;
            }

            catch (Exception EX)
            {
                etat = false;
            }
            finally
            {
                if (con != null)
                {
                    con.Close();
                    con = null;
                }
            }

            return etat;
        }

        public bool insertGrossToNetA014(string IdOracle, string Matricule, string Nom, string Prenom, string DateEmbauche, string DateDepart,string LocalExpat, string Emploi,
           string IntituleDept,string CentreCout,string A1000,string A1910,string A1920,string A3310,string A3317,string A3711,string HS01,string A4110,string HS02,string A4111,
            string HS03,string A4113,string A4115,string A4100,string A4170,string A4159,string A4175,string A4224,string A4411,string A4816,string A4817,string A4222,string A4340,string A4360,
            string A4361,string A4380,string A4381,string A4410,string A4412,string A4450,string A4460,string A4550,string A4711,string A4712,string A4715,string A4716,string A4717,
            string A4790,string A4811,string A4812,string A4813,string IFRT,string IndemRet,string Rappel,string A5420,string A5200,string PrimeTrim,string CL03,string A4171,string A7193,
            string A7114,string A7117,string A7191,string A7192,string A7197,string A7196,string A7198,string A7199,string A7692,string A7115,string A7118,string BRUT,string A9112,
            string A6110,string A6120,string A6201,string A6221,string A6222,string A9121,string A6310,string A6311,string A7115A,string A7118A,string A7193A,string A7114A,string A6837,string A6832,
            string A7197A,string A6842,string A6843,string A6844,string A6850,string A6410,string A6430,string A6470,string A6491,string A6614,string A6530,string A6510,string A6521,
            string A6520,string A8201,string A8215,string A8217,string RUB_DEDSA,string NETPAIE,string A6110A,string A6120A,string A6150,string A7115AA,string A7118AA,string RUB_COTEPY,
            string MoisPe, int IdSoc, int Mois, int Annee,string A5400,string A4610)
        {
            OleDbConnection con = null;
            bool etat = false;
            try
            {
                con = new OleDbConnection(connection_string);
                con.Open();
                string select = "insert into GROSSTONETA014(IdOracle,Matricule,Nom,Prenom,DateEmbauche,DateDepart,LocalExpat,Emploi,IntituleDept,CentreCout,A1000,A1910,"
                    +"A1920,A3310,A3317,A3711,HS01,A4110,HS02,A4111,HS03,A4113,A4115,A4100,A4170,A4159,A4175,A4224,A4411,A4816,A4817,A4222,A4340,A4360,A4361,A4380,A4381,"
                    +"A4410,A4412,A4450,A4460,A4550,A4711,A4712,A4715,A4716,A4717,A4790,A4811,A4812,A4813,IFRT,IndemRet,A5100,A5420,A5200,PrimeTrim,CL03,A4171,A7193,A7114,"
                    + "A7117,A7191,A7192,A7197,A7196,A7198,A7199,A7692,A7115,A7118,BRUT,A9112,A6110,A6120,A6201,A6221,A6222,A9121,A6310,A6311,A6840,A6841,A6835,A6831,A6837,"
                    +"A6832,A6834,A6842,A6843,A6844,A6850,A6410,A6430,A6470,A6491,A6614,A6530,A6510,A6521,A6520,A8201,A8215,A8217,RUB_DEDSA,NETPAIE,A61109,A61209,A6150,"
                    +"A711599,A711899,RUB_COTEPY,MoisPe,IdSoc,Mois,Annee,A5400,A4610)"
               + " values('" + IdOracle + "','" + Matricule + "','" + Nom + "','" + Prenom + "','" + DateEmbauche + "','" + DateDepart + "','" + LocalExpat + "','" + Emploi + "'"
                + ",'" + IntituleDept + "','" + CentreCout + "','" + A1000 + "','" + A1910 + "','" + A1920 + "','" + A3310 + "' ,'" + A3317 + "'"
                + ",'" + A3711 + "','" + HS01 + "','" + A4110 + "','" + HS02 + "','" + A4111 + "','" + HS03 + "','" + A4113 + "'"
                + ",'" + A4115 + "','" + A4100 + "','" + A4170 + "','" + A4159 + "','" + A4175 + "','" + A4224 + "','" + A4411 + "','" + A4816 + "','"+A4817+"','" + A4222 + "'"
                + ",'" + A4340 + "','" + A4360 + "','" + A4361 + "','" + A4380 + "','" + A4381 + "','" + A4410 + "'"
                + ",'" + A4412 + "','" + A4450 + "','" + A4460 + "','" + A4550 + "','" + A4711 + "','" + A4712 + "','" + A4715 + "'"
                + ",'" + A4716 + "','" + A4717 + "','" + A4790 + "','" + A4811 + "','" + A4812 + "','" + A4813 + "','" + IFRT + "'"
                + ",'" + IndemRet + "','" + Rappel + "','" + A5420 + "','"+A5200+"','" + PrimeTrim + "','" + CL03 + "','" + A4171 + "','" + A7193 + "','" + A7114 + "','" + A7117 + "'"
                + ",'" + A7191 + "','" + A7192 + "','" + A7197 + "','" + A7196 + "','" + A7198 + "','" + A7199 + "','" + A7692 + "','" + A7115 + "'"
                + ",'" + A7118 + "','" + BRUT + "','" + A9112 + "','" + A6110 + "','" + A6120 + "','" + A6201 + "','" + A6221 + "','"+A6222+"','" + A9121 + "','" + A6310 + "'"
                + ",'" + A6311 + "','" + A7115A + "','" + A7118A + "','" + A7193A + "','" + A7114A + "','" + A6837 + "','" + A6832 + "','" + A7197A + "','" + A6842 + "'"
                + ",'" + A6843 + "','" + A6844 + "','" + A6850 + "','" + A6410 + "','" + A6430 + "','" + A6470 + "','" + A6491 + "','" + A6614 + "','" + A6530 + "','" + A6510 + "'"
                +",'"+A6521+"','"+A6520+"','"+A8201+"','"+A8215+"','"+A8217+"','"+RUB_DEDSA+"','"+NETPAIE+"','"+A6110A+"','"+A6120A+"','"+A6150+"','"+A7115AA+"','"+A7118AA+"'"
                + ",'" + RUB_COTEPY + "','" + MoisPe + "'," + IdSoc + "," + Mois + "," + Annee + ",'"+A5400+"','"+A4610+"')";
                OleDbCommand cmd = new OleDbCommand(select, con);
                cmd.ExecuteNonQuery();
                etat = true;
            }

            catch (Exception ex)
            {
                etat = false;
            }
            finally
            {
                if (con != null)
                {
                    con.Close();
                    con = null;
                }
            }

            return etat;
        }

        public bool insertGrossToNetA016(string Matricule, string Nom, string Prenom, string DateEmbauche, string DateDepart, string CodeDept, string IntituleDept,
          string NatureContrat, string Emploi, string SALBASE, string HORAIRE, string A1000, string A2100, string A2200, string A3413, string A3414, string A3415, string A3418,
          string A3316, string A3317, string A3510, string A3520, string A3530, string A3417, string A5910, string A3318, string A3610, string A3620, string A3630, string A3640, string A3651,
          string A3652, string A3650, string A3660, string A3661, string A3670, string A3680, string A3681, string A3682, string A3683, string A3684, string A3685, string BRUTFX, string A3686,
          string A3690, string CL31, string A4460, string CL32, string A4461, string CL33, string A4462, string CL34, string A4463, string CL35, string A3110,
          string CL37, string A4292, string A4293, string CL39, string A4294, string HC01, string A4295, string CL40, string A4296, string A4491, string A4162, string A4391, string A4381,
          string A4840, string A4382, string A4850, string A4383, string A4860, string A4718, string A4717, string A4713, string A4716, string HS01, string A4111, string HS02, string A4112,
          string HS03, string A4113, string HS04, string A4114, string HS05, string A4115, string HS06, string A4116, string HS07, string A4117, string A4172, string A4173,
          string HC02, string A4464, string A5100, string A5700, string A5702, string A5220, string A5400, string A5800, string A5900, string A7111, string A7114, string A7112,
          string A7116, string A7117, string A7113, string A5920, string A59209, string A7115, string A7600, string A7670, string A3910, string A4170, string A4811, string A4813, string A4812,
          string A4830, string BRUT, string A9110, string A9112, string A6110, string A6120, string A9121, string A6310, string A6312, string A6311, string A6462, string A6511, string A6430,
          string A6465, string A6467, string A6840, string A6810, string A6830, string A6860, string A6463, string A6610, string A6870, string A8140, string RUB_DEDSA, string NETPAIE,
          string A61109, string A61209, string A6150, string A6210, string TFP, string RUB_COTEPY, string RD_NET, string MoisPe, int IdSoc, int Mois, int Annee, string RIB, string CUMNB2)
        {
            OleDbConnection con = null;
            bool etat = false;
            try
            {
                con = new OleDbConnection(connection_string);
                con.Open();
                string select = "insert into GROSSTONETA016( Matricule, Nom, Prenom, DateEmbauche, DateDept, CodeDept, IntituleDept,"
                +"NatureContrat,  Emploi,  SALBASE,  HORAIRE,  A1000,  A2100,  A2200,  A3413,  A3414,  A3415,  A3418,"
                +" A3416,  A3317,  A3510,  A3520,  A3530,  A3417,  A5910,  A3318,  A3610,  A3620,  A3630,  A3640,  A3651,"
                +"A3652,  A3650,  A3660,  A3661,  A3670,  A3680,  A3681,  A3682,  A3683,  A3684,  A3685,  BRUTFX,  A3686,"
                +"A3690,  CL31,  A4460,  CL32,  A4461,  CL33,  A4462,  CL34,  A4463,  CL35,  A3110,"
                +"CL37,  A4292,  A4293,  CL39,  A4294,  HC01,  A4295,  CL40,  A4296,  A4491,  A4162,  A4391,  A4381,"
                +"A4840,  A4382,  A4850,  A4383,  A4860,  A4718,  A4717,  A4713,  A4716,  HS01,  A4111,  HS02,  A4112,"
                +"HS03,  A4113,  HS04,  A4114,  HS05,  A4115,  HS06,  A4116,  HS07,  A4117,  A4172,  A4173,"
                +"HC02,  A4464,  A5100,  A5700,  A5702,  A5220,  A5400,  A5800,  A5900,  A7111,  A7114,  A7112,"
                +"A7116,  A7117,  A7113,  A5920,  A59209,  A7115,  A7600,  A7670,  A3910,  A4170,  A4811,  A4813,  A4812,"
               +" A4830,  BRUT,  A9110,  A9112,  A6110,  A6120,  A9121,  A6310,  A6312,  A6311,  A6462,  A6511,  A6430,"
               +" A6465,  A6467,  A6840,  A6810,  A6830,  A6860,  A6463,  A6610,  A6870,  A8140,  RUB_DEDSA,  NETPAIE,"
               + "A61109,  A61209,  A6150,  A6210,  TFP,  RUB_COTEPY,  RD_NET,  MoisPe,  IdSoc,  Mois,  Annee,RIB,CUMNB2)"
               + " values('" + Matricule + "','" + Nom + "','" + Prenom + "','" + DateEmbauche + "','" + DateDepart + "','" + CodeDept + "','" + IntituleDept + "'"
                + ",'" + NatureContrat + "','" + Emploi + "','" + SALBASE + "','" + HORAIRE + "','" + A1000 + "','" + A2100 + "' ,'" + A2200 + "'"
                + ",'" + A3413 + "','" + A3414 + "','" + A3415 + "','" + A3418 + "','" + A3316 + "','" + A3317 + "','" + A3510 + "'"
                + ",'" + A3520 + "','" + A3530 + "','" + A3417 + "','" + A5910 + "','" + A3318 + "','" + A3610 + "','" + A3620 + "','" + A3630 + "','" + A3640 + "','" + A3651 + "'"
                + ",'" + A3652 + "','" + A3650 + "','" + A3660 + "','" + A3661 + "','" + A3670 + "','" + A3680 + "'"
                + ",'" + A3681 + "','" + A3682 + "','" + A3683 + "','" + A3684 + "','" + A3685 + "','" + BRUTFX + "','" + A3686 + "'"
                + ",'" + A3690 + "','" + CL31 + "','" + A4460 + "','" + CL32 + "','" + A4461 + "','" + CL33 + "','" + A4462 + "'"
                + ",'" + CL34 + "','" + A4463 + "','" + CL35 + "','" + A3110 + "','" + CL37 + "','" + A4292 + "','" + A4293 + "','" + CL39 + "','" + A4294 + "','" + HC01 + "'"
                + ",'" + A4295 + "','" + CL40 + "','" + A4296 + "','" + A4491 + "','" + A4162 + "','" + A4391 + "','" + A4381 + "','" + A4840 + "'"
                + ",'" + A4382 + "','" + A4850 + "','" + A4383 + "','" + A4860 + "','" + A4718 + "','" + A4717 + "','" + A4713 + "','" + A4716 + "','" + HS01 + "','" + A4111 + "'"
                + ",'" + HS02 + "','" + A4112 + "','" + HS03 + "','" + A4113 + "','" + HS04 + "','" + A4114 + "','" + HS05 + "','" + A4115 + "','" + HS06 + "'"
                + ",'" + A4116 + "','" + HS07 + "','" + A4117 + "','" + A4172 + "','" + A4173 + "','" + HC02 + "','" + A4464 + "','" + A5100 + "','" + A5700 + "','" + A5702 + "'"
                + ",'" + A5220 + "','" + A5400 + "','" + A5800 + "','" + A5900 + "','" + A7111 + "','" + A7114 + "','" + A7112 + "','" + A7116 + "','" + A7117 + "','" + A7113 + "','" + A5920 + "','" + A59209 + "'"
                + ",'" + A7115 + "','" + A7600 + "','" + A7670 + "','" + A3910 + "','" + A4170 + "','" + A4811 + "','" + A4813 + "','" + A4812 + "','" + A4830 + "','" + BRUT + "'"
                + ",'" + A9110 + "','" + A9112 + "','" + A6110 + "','" + A6120 + "','" + A9121 + "','" + A6310 + "','" + A6312 + "','" + A6311 + "','" + A6462 + "','" + A6511 + "','" + A6430 + "','" + A6465 + "'"
                + ",'" + A6467 + "','" + A6840 + "','" + A6810 + "','" + A6830 + "','" + A6860 + "','" + A6463 + "','" + A6610 + "','" + A6870 + "','" + A8140 + "','" + RUB_DEDSA + "','" + NETPAIE + "','" + A61109 + "'"
                + ",'" + A61209 + "','" + A6150 + "','" + A6210 + "','" + TFP + "','" + RUB_COTEPY + "','" + RD_NET + "','" + MoisPe + "'," + IdSoc + "," + Mois + "," + Annee + ",'" + RIB + "','" + CUMNB2 + "')";
                OleDbCommand cmd = new OleDbCommand(select, con);
                cmd.ExecuteNonQuery();
                etat = true;
            }

            catch (Exception ex)
            {
                etat = false;
            }
            finally
            {
                if (con != null)
                {
                    con.Close();
                    con = null;
                }
            }

            return etat;
        }

        public bool insertGrossToNetA020(string IdOracle, string Matricule, string Nom, string Prenom, string DateEmbauche, string DateDepart, string CodeDept,string IntituleDept,
        string ExpatLocal,string Emploi,string CentreCout,string A1000,string A1910,string A1920,string A2311,string A3317,string A3510,string A3711,string HS01,string A4110,
        string HS02,string A4111,string HS03,string A4113,string A4115,string A4100,string A4159,string A4175,string A4224,string A4160,string A4222,string A4340,string A4360,
            string A4361,string A4380,string A4381,string A7117,string A4410,string A4816,string A4817,string A4412,string A4450,string A4460,string A4550,string A4711,string A4712,
            string A4715,string A4716,string A4717,string A4790,string A4811,string A4812,string A4813,string IFRT,string IndemRetAnt,string LiquidIfrt,string A5100,
            string A5420,string A5200,string PrimTrim,string CongePris,string A4171,string HA04,string A4170,string A7114,string A7191,string A7193,string A7196,string A7692,string A7115,
            string A7118,string BRUT,string A9112,string A6110,string A6120,string A6201,string A6221,string A9121,string A6310,string A6311,string A6840,string A6841,string A6835,
            string A6831,string A44509,string A6832,string A6834,string A6850,string A6410,string A6430,string A6470,string A6491,string A6520,string A6510,string A6521,
            string A6591,string A6614,string A8201,string A8215,string A8217,string RUB_DEDSA,string NETPAIE,string A61109,string A61209,string A6150,string A71159,string A62019,
            string RUB_COTEPY,string MoisPe, int IdSoc, int Mois, int Annee,string A5400)
        {
            OleDbConnection con = null;
            bool etat = false;
            try
            {
                con = new OleDbConnection(connection_string);
                con.Open();
                string select = "insert into GROSSTONETA020(IdOracle,Matricule,Nom,Prenom,DateEmbauche,DateDepart,CodeDept,IntituleDept,ExpatLocal,Emploi,CentreCout,A1100,"
                +"A1910,A1920,A2311,A3317,A3510,A3711,HS01,A4110,HS02,A4111,HS03,A4113,A4115,A4100,A4159,A4175,A4224,A4160,A4222,A4340,A4360,A4361,A4380,A4381,A7117,A4410,A4816,A4817,A4412,A4450,A4460,"
                + "A4550,A4711,A4712,A4715,A4716,A4717,A4790,A4811,A4812,A4813,IFRT,IndemRetAnt,LiquidIfrt,A5100,A5420,A5200,PrimTrim,COPRIMOI,A4171,HA04,A4170,A7114,A7191,A7193,A7196,A7692,A7115,"
                +"A7118,BRUT,A9112,A6110,A6120,A6201,A6221,A9121,A6310,A6311,A6840,A6841,A6835,A6831,A44509,A6832,A6834,A6850,A6410,A6430,A6470,A6491,A6520,A6510,A6521,A6591,A6614,A8201,A8215,A8217,"
                +"RUB_DEDSA,NETPAIE,A61109,A61209,A6150,A71159,A62019,RUB_COTEPY,MoisPe,IdSoc,Mois,Annee,A5400)"
               + " values('" + IdOracle + "','" + Matricule + "','" + Nom + "','" + Prenom + "','" + DateEmbauche + "','" + DateDepart + "','" + CodeDept + "'"
                + ",'" + IntituleDept + "','" + ExpatLocal + "','" + Emploi + "','" + CentreCout + "','" + A1000 + "','" + A1910 + "' ,'" + A1920 + "'"
                + ",'" + A2311 + "','" + A3317 + "','" + A3510 + "','" + A3711 + "','" + HS01 + "','" + A4110 + "','" + HS02 + "'"
                + ",'" + A4111 + "','" + HS03 + "','" + A4113 + "','" + A4115 + "','" + A4100 + "','" + A4159 + "','" + A4175 + "','" + A4224 + "'"
                + ",'" + A4160 + "','" + A4222 + "','" + A4340 + "','" + A4360 + "','" + A4361 + "','" + A4380 + "'"
                + ",'" + A4381 + "','" + A7117 + "','" + A4410 + "','" + A4816 + "','"+A4817+"','" + A4412 + "','" + A4450 + "','" + A4460 + "'"
                + ",'" + A4550 + "','" + A4711 + "','" + A4712 + "','" + A4715 + "','" + A4716 + "','" + A4717 + "','" + A4790 + "'"
                + ",'" + A4811 + "','" + A4812 + "','" + A4813 + "','" + IFRT + "','" + IndemRetAnt + "','" + LiquidIfrt + "','" + A5100 + "','" + A5420 + "','"+A5200+"','" + PrimTrim + "'"
                + ",'" + CongePris + "','" + A4171 + "','" + HA04 + "','" + A4170 + "','" + A7114 + "','" + A7191 + "','" + A7193 + "','" + A7196 + "'"
                + ",'" + A7692 + "','" + A7115 + "','" + A7118 + "','" + BRUT + "','" + A9112 + "','" + A6110 + "','" + A6120 + "','" + A6201 + "','" + A6221 + "'"
                + ",'" + A9121 + "','" + A6310 + "','" + A6311 + "','" + A6840 + "','" + A6841 + "','" + A6835 + "','" + A6831 + "','" + A44509 + "','" + A6832 + "'"
                + ",'" + A6834 + "','" + A6850 + "','" + A6410 + "','" + A6430 + "','" + A6470 + "','" + A6491 + "','" + A6520 + "'"
                + ",'" + A6510 + "','" + A6521 + "','" + A6591 + "','" + A6614 + "','" + A8201 + "','" + A8215 + "','" + A8217 + "','" + RUB_DEDSA + "','" + NETPAIE + "','" + A61109 + "','" + A61209 + "','" + A6150 + "'"
                + ",'" + A71159 + "','"+A62019+"','"+RUB_COTEPY+"','" + MoisPe + "'," + IdSoc + "," + Mois + "," + Annee + ",'"+A5400+"')";
                OleDbCommand cmd = new OleDbCommand(select, con);
                cmd.ExecuteNonQuery();
                etat = true;
            }

            catch (Exception ex)
            {
                etat = false;
            }
            finally
            {
                if (con != null)
                {
                    con.Close();
                    con = null;
                }
            }

            return etat;
        }
    }
}
