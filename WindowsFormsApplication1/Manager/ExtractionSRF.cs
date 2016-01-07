using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using ExcelLibrary.SpreadSheet;
using System.Data;
using System.Configuration;

namespace HarmonizedApp.Manager
{
    class ExtractionSRF
    {
         
        ServiceDAO service = new ServiceDAO();
        EnteteSRF entete = new EnteteSRF();
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

        public void EcrireDataA001(int NumSoc,string file,ProgressBar prog)
        {
            StreamReader sr = new StreamReader(file);
            string line = null;
            prog.Visible = true;
            prog.Minimum = 0;     
            int ligne = 0;
            string matricule=string.Empty;
            string nom = string.Empty; 
            string prenom = string.Empty;
            while ((line = sr.ReadLine()) != null)
            {
                ligne = ligne + 1;
                prog.Maximum = ligne;
                if (!(line.Equals("")))
                {                  
                         matricule = line.Substring(0, 10).Trim();
                         nom = line.Substring(10, 80).Trim().Replace("'", "''");
                         prenom = line.Substring(90, 20).Trim().Replace("'", "''");
                         if ((matricule != "") && (nom != "") && (prenom != ""))
                         {
                            string DateEmbauche = line.Substring(110, 8).Trim();
                            string DateDept = line.Substring(118, 8).Trim();
                            string CodeSce = line.Substring(126, 10).Trim();
                            string IntituleSce = line.Substring(136, 30).Trim().Replace("'", "''");
                            string NatureContrat = line.Substring(166, 10).Trim();
                            string Emploi = line.Substring(176, 30).Trim().Replace("'", "''");
                            string A1100 = line.Substring(206, 12).Trim();
                            string CL01 = line.Substring(218, 12).Trim();
                            string A1200 = line.Substring(230, 12).Trim();
                            string A2200 = line.Substring(242, 12).Trim();
                            string A2100 = line.Substring(254, 12).Trim();
                            string A2311 = line.Substring(266, 12).Trim();
                            string A2320 = line.Substring(278, 12).Trim();
                            string A3318 = line.Substring(290, 12).Trim();
                            string A3110 = line.Substring(302, 12).Trim();
                            string A4311 = line.Substring(314, 12).Trim();
                            string A4330 = line.Substring(326, 12).Trim();
                            string A3111 = line.Substring(338, 12).Trim();
                            string A3210 = line.Substring(350, 12).Trim();
                            string A3317 = line.Substring(362, 12).Trim();
                            string A3319 = line.Substring(374, 12).Trim();
                            string A4100 = line.Substring(386, 12).Trim();
                            string A4171 = line.Substring(398, 12).Trim();
                            string A4130 = line.Substring(410, 12).Trim();
                            string A4223 = line.Substring(422, 12).Trim();
                            string A4380 = line.Substring(434, 12).Trim();
                            string A4711 = line.Substring(446, 12).Trim();
                            string A4717 = line.Substring(458, 12).Trim();
                            string A5100 = line.Substring(470, 12).Trim();
                            string A7600 = line.Substring(482, 12).Trim();
                            string A4811 = line.Substring(494, 12).Trim();
                            string A4812 = line.Substring(506, 12).Trim();
                            string A4813 = line.Substring(518, 12).Trim();
                            string BRUT = line.Substring(530, 12).Trim();
                            string AjusCot = line.Substring(542, 12).Trim();
                            string A9112 = line.Substring(554, 12).Trim();
                            string A6110 = line.Substring(566, 12).Trim();
                            string A6120 = line.Substring(578, 12).Trim();
                            string A6210 = line.Substring(590, 12).Trim();
                            string A6230 = line.Substring(602, 12).Trim();
                            string A9121 = line.Substring(614, 12).Trim();
                            string A6310 = line.Substring(626, 12).Trim();
                            string A6311 = line.Substring(638, 12).Trim();
                            string A6614 = line.Substring(650, 12).Trim();
                            string A6410 = line.Substring(662, 12).Trim();
                            string A6520 = line.Substring(674, 12).Trim();
                            string A6510 = line.Substring(686, 12).Trim();
                            string A6530 = line.Substring(698, 12).Trim();
                            string A6890 = line.Substring(710, 12).Trim();
                            string A8000 = line.Substring(722, 12).Trim();
                            string RUB_DEDSA = line.Substring(734, 12).Trim();
                            string NETPAIE = line.Substring(746, 12).Trim();
                            string A61109 = line.Substring(758, 12).Trim();
                            string A61209 = line.Substring(770, 12).Trim();
                            string A6150 = line.Substring(782, 12).Trim();
                            string A62109 = line.Substring(794, 12).Trim();
                            string A62309 = line.Substring(806, 12).Trim();
                            string A6320 = line.Substring(818, 12).Trim();
                            string A6330 = line.Substring(830, 12).Trim();
                            string RUB_COTEPY = line.Substring(842, 12).Trim();
                            string A6490 = line.Substring(854, 12).Trim();
                            string MoisPe = line.Substring(866, 9).Trim();
                            string MoisInt = line.Substring(875, 2).Trim();
                            string A4712 = line.Substring(877, 12).Trim();
                            string A4226 = line.Substring(889, 12).Trim();
                            prog.Value = prog.Value + 1;
                            int Annee = DateTime.Today.Year;
                            int Mois = int.Parse(MoisInt);
                       
                            service.insertGrossToNetA001(matricule, nom.Replace("�", "e"), prenom.Replace("�", "e"), DateEmbauche, DateDept, CodeSce, IntituleSce.Replace("�", "e"),
                                 NatureContrat, Emploi.Replace("�", "e"), A1100, CL01, A1200, A2200, A2100, A2311,
                                 A2320, A3318, A3110, A4311, A4330, A3111, A3210, A3317, A3319, A4100,
                                 A4171,A4130, A4223, A4380, A4711, A4717, A5100, A7600, A4811, A4812, A4813, BRUT, AjusCot,
                                A9112, A6110, A6120, A6210, A6230, A9121, A6310, A6311, A6614, A6410, A6520,A6510, A6530,
                                 A6890,A8000, RUB_DEDSA, NETPAIE, A61109, A61209, A6150, A62109, A62309, A6320, A6330, RUB_COTEPY,A6490,
                                 MoisPe.Replace("�", "e"), NumSoc.ToString(), Mois, Annee,A4712,A4226);
                    }
                }

            }
            MessageBox.Show("Import realized with success");
     //       prog.Value = 0;
     //       prog.Visible = false;*/
        }

        public void EcrireDataA002(int NumSoc, string file, ProgressBar prog)
        {
            StreamReader sr = new StreamReader(file);
            string line = null;
            prog.Visible = true;
            prog.Minimum = 0;
            int ligne = 0;
            string matricule = string.Empty;
            string nom = string.Empty;
            string prenom = string.Empty;
            while ((line = sr.ReadLine()) != null)
            {
                ligne = ligne + 1;
                prog.Maximum = ligne;
                if (!(line.Equals("")))
                {                   
                         matricule = line.Substring(0, 10).Trim();
                         nom = line.Substring(10, 80).Trim().Replace("'", "''");
                         prenom = line.Substring(90, 20).Trim().Replace("'", "''");
                         if ((matricule != "") && (nom != "") && (prenom != ""))
                         {
                            string DateEmbauche = line.Substring(110, 8).Trim();
                            string DateDept = line.Substring(118, 8).Trim();
                            string CodeSce = line.Substring(126, 10).Trim();
                            string IntituleSce = line.Substring(136, 30).Trim().Replace("'", "''");
                            string NatureContrat = line.Substring(166, 10).Trim();
                            string Emploi = line.Substring(176, 30).Trim().Replace("'", "''");
                            string SalBase = line.Substring(206, 12).Trim();
                            string NbJrsPresence = line.Substring(218, 12).Trim();
                            string NbHrPresence = line.Substring(230, 12).Trim();
                            string SalBaseJrsTrav = line.Substring(242, 12).Trim();
                            string IndemTrans = line.Substring(254, 12).Trim();
                            string IndemPre = line.Substring(266, 12).Trim();
                            string CongePaye = line.Substring(278, 12).Trim();
                            string CongeSpecio = line.Substring(290, 12).Trim();
                            string CongeAbcnonRem = line.Substring(302, 12).Trim();
                            string PrimeExcepFixe = line.Substring(314, 12).Trim();
                            string CongeAPayer = line.Substring(326, 12).Trim();
                            string HS125 = line.Substring(338, 12).Trim();
                            string HS150 = line.Substring(350, 12).Trim();
                            string HS200 = line.Substring(362, 12).Trim();
                            string PrimeAstreinte = line.Substring(374, 12).Trim();
                            string PrimeDim = line.Substring(386, 12).Trim();
                            string PrimeDiff = line.Substring(398, 12).Trim();
                            string PrimDifFormation = line.Substring(410, 12).Trim();
                            string PrimeParrain = line.Substring(422, 12).Trim();
                            string PrimeFormation = line.Substring(434, 12).Trim();
                            string PrimeExcep = line.Substring(446, 12).Trim();
                            string PrimeObj = line.Substring(458, 12).Trim();
                            string PrimeNaiss = line.Substring(470, 12).Trim();
                            string PrimeMariage = line.Substring(482, 12).Trim();
                            string PrimeDeces = line.Substring(494, 12).Trim();
                            string PrimAid = line.Substring(506, 12).Trim();
                            string RappelEltRec = line.Substring(518, 12).Trim();
                            string RappelEltnonRec = line.Substring(530, 12).Trim();
                            string AvantageEnNatTR = line.Substring(542, 12).Trim();
                            string RappelAvNatTr = line.Substring(554, 12).Trim();
                            string AvantageEnNatAssGr = line.Substring(566, 12).Trim();
                            string IndemPreavis = line.Substring(578, 12).Trim();
                            string GratifFinSce = line.Substring(590, 12).Trim();
                            string IndemLic = line.Substring(602, 12).Trim();
                            string A4815 = line.Substring(614, 12).Trim();

                            string BRUT = line.Substring(626, 12).Trim();

                            string A9112 = line.Substring(638, 12).Trim();

                            string A6110 = line.Substring(650, 12).Trim();

                            string A6210 = line.Substring(662, 12).Trim();

                            string A9121 = line.Substring(674, 12).Trim();

                            string A6310 = line.Substring(686, 12).Trim();

                            string A6311 = line.Substring(698, 12).Trim();

                            string A6614 = line.Substring(710, 12).Trim();

                            string A6410 = line.Substring(722, 12).Trim();

                            string A6520 = line.Substring(734, 12).Trim();

                            string A6430 = line.Substring(746, 12).Trim();

                            string A6490 = line.Substring(758, 12).Trim();

                            string A6830 = line.Substring(770, 12).Trim();

                            string A6831 = line.Substring(782, 12).Trim();

                            string A6840 = line.Substring(794, 12).Trim();

                            string A8800 = line.Substring(806, 12).Trim();

                            string RUB_DEDSA = line.Substring(818, 12).Trim();

                            string NETPAIE = line.Substring(830, 12).Trim();

                            string A61109 = line.Substring(842, 12).Trim();

                            string A6150 = line.Substring(854, 12).Trim();

                            string A71159 = line.Substring(866, 12).Trim();

                            string RUB_COTEPY = line.Substring(878, 12).Trim();

                            string MoisPe = line.Substring(890, 9).Trim();

                            string MoisEnt = line.Substring(899, 2).Trim();

                            string A5100 = line.Substring(901, 12).Trim();

                            string A7116 = line.Substring(913, 12).Trim();

                            string A6850 = line.Substring(925, 12).Trim();

                            prog.Value = prog.Value + 1;
                            int Annee = DateTime.Today.Year;
                            int Mois = int.Parse(MoisEnt);                  
                            service.insertGrossToNetA002(matricule, nom.Replace("�", "e"), prenom.Replace("�", "e"), DateEmbauche, DateDept, CodeSce, IntituleSce.Replace("�", "e"),
                                 NatureContrat, Emploi.Replace("�", "e"), SalBase, NbJrsPresence, NbHrPresence, SalBaseJrsTrav, IndemTrans, IndemPre,CongePaye,CongeSpecio,
                                 CongeAbcnonRem, PrimeExcepFixe, CongeAPayer, HS125, HS150, HS200, PrimeAstreinte, PrimeDim, PrimeDiff, PrimDifFormation,PrimeParrain,PrimeFormation, PrimeExcep, PrimeObj,
                                 PrimeNaiss, PrimeMariage, PrimeDeces, PrimAid, RappelEltRec, RappelEltnonRec, AvantageEnNatTR,RappelAvNatTr, AvantageEnNatAssGr, IndemPreavis, GratifFinSce, IndemLic,
                                A4815, BRUT, A9112, A6110, A6210, A9121, A6310, A6311, A6614, A6410, A6520, A6430,A6490,
                                 A6830, A6831, A6840, A8800, RUB_DEDSA, NETPAIE, A61109, A6150, A71159,RUB_COTEPY,
                                 MoisPe.Replace("�", "e"), NumSoc.ToString(), Mois, Annee,A5100,A7116,A6850);
                    }
                }

            }
            MessageBox.Show("Import realized with success");
            //       prog.Value = 0;
            //       prog.Visible = false;*/
        }

        public void EcrireDataA003(int NumSoc, string file, ProgressBar prog)
        {
            StreamReader sr = new StreamReader(file);
            string line = null;
            prog.Visible = true;
            prog.Minimum = 0;
            int ligne = 0;
            string matricule = string.Empty;
            string nom = string.Empty;
            string prenom = string.Empty;
            while ((line = sr.ReadLine()) != null)
            {
                ligne = ligne + 1;
                prog.Maximum = ligne;
                if (!(line.Equals("")))
                {
                    matricule = line.Substring(0, 10).Trim();
                    nom = line.Substring(10, 80).Trim().Replace("'", "''");
                    prenom = line.Substring(90, 20).Trim().Replace("'", "''");
                    if ((matricule != "") && (nom != "") && (prenom != ""))
                    {
                        string DateEmbauche = line.Substring(110, 8).Trim();
                        string DateDept = line.Substring(118, 8).Trim();
                        string CodeSce = line.Substring(126, 10).Trim();
                        string IntituleSce = line.Substring(136, 30).Trim().Replace("'", "''");
                        string NatureContrat = line.Substring(166, 10).Trim();
                        string Emploi = line.Substring(176, 30).Trim().Replace("'", "''");
                        string SalBase = line.Substring(206, 12).Trim();
                        string CL01 = line.Substring(218, 12).Trim();
                        string A1200 = line.Substring(230, 12).Trim();
                        string A2200 = line.Substring(242, 12).Trim();
                        string A2100 = line.Substring(254, 12).Trim();
                        string A2360 = line.Substring(266, 12).Trim();
                        string A2330 = line.Substring(278, 12).Trim();
                        string A3412 = line.Substring(290, 12).Trim();
                        string A3317 = line.Substring(302, 12).Trim();
                        string A3712 = line.Substring(314, 12).Trim();
                        string A3710 = line.Substring(326, 12).Trim();
                        string A3711 = line.Substring(338, 12).Trim();
                        string A3713 = line.Substring(350, 12).Trim();
                        string A4112 = line.Substring(362, 12).Trim();
                        string A4113 = line.Substring(374, 12).Trim();
                        string  A4120= line.Substring(386, 12).Trim();
                        string  A4222= line.Substring(398, 12).Trim();
                        string A4216 = line.Substring(410, 12).Trim();
                        string A4361 = line.Substring(422, 12).Trim();
                        string A4215 = line.Substring(434, 12).Trim();
                        string A4320 = line.Substring(446, 12).Trim();

                        string A4370 = line.Substring(458, 12).Trim();

                      //  string A4380 = line.Substring(470, 12).Trim();

                        string A4430 = line.Substring(470, 12).Trim();

                        string A7113 = line.Substring(482, 12).Trim();

                        string A4171 = line.Substring(494, 12).Trim();

                        string BRUT = line.Substring(506, 12).Trim();

                        string AjusCot = line.Substring(518, 12).Trim();

                        string A9112 = line.Substring(530, 12).Trim();

                        string A6110 = line.Substring(542, 12).Trim();

                        string A6120 = line.Substring(554, 12).Trim();

                        string A9310 = line.Substring(566, 12).Trim();

                        string A6310 = line.Substring(578, 12).Trim();

                        string A6311 = line.Substring(590, 12).Trim();

                        string A6340 = line.Substring(602, 12).Trim();

                        string A6210 = line.Substring(614, 12).Trim();

                        string A6410 = line.Substring(626, 12).Trim();

                        string A6430 = line.Substring(638, 12).Trim();

                        string A6520 = line.Substring(650, 12).Trim();

                        string A6510 = line.Substring(662, 12).Trim();

                        string A6530 = line.Substring(674, 12).Trim();

                        string A6830 = line.Substring(686, 12).Trim();

                        string A6490 = line.Substring(698, 12).Trim();

                        string RUB_DEDSA = line.Substring(710, 12).Trim();

                        string NETPAIE = line.Substring(722, 12).Trim();

                        string A6110A = line.Substring(734, 12).Trim();

                        string A6120A = line.Substring(746, 12).Trim();

                        string A6150 = line.Substring(758, 12).Trim();

                        string A6210A = line.Substring(770, 12).Trim();

                        string RUB_COTEPY = line.Substring(782, 12).Trim();

                        string MoisPe = line.Substring(794, 9).Trim();

                        string MoisEnt = line.Substring(803, 2).Trim();

                        string A4812 = line.Substring(813, 4).Trim();

                        string A4380 = line.Substring(821, 9).Trim();

                        string A5100 = line.Substring(836, 5).Trim();

                      

                        prog.Value = prog.Value + 1;
                        int Annee = 2015;//DateTime.Today.Year;
                        int Mois = 12;// int.Parse(MoisEnt);
                        service.insertGrossToNetA003(matricule, nom.Replace("�", "e"), prenom.Replace("�", "e"), DateEmbauche,
                            DateDept, CodeSce, IntituleSce.Replace("�", "e"), NatureContrat, Emploi.Replace("�", "e"), SalBase,
                              CL01, A1200, A2200, A2100, A2360, A2330, A3412, A3317, A3712, A3710, A3711, A3713,A4112,
                             A4113, A4120, A4222,A4216, A4361, A4215, A4320, A4370, A4430, A7113, A4171, BRUT, AjusCot, A9112, A6110,
                              A6120, A9310, A6310, A6311, A6340, A6210, A6410, A6430, A6520, A6510, A6530,
                             A6830,A6490, RUB_DEDSA, NETPAIE, A6110A, A6120A, A6150, A6210A, RUB_COTEPY, MoisPe.Replace("�", "e"),
                              NumSoc.ToString(), Mois, Annee,A4812,A5100,A4380);
                    }
                }

            }
            MessageBox.Show("Import realized with success");
            //       prog.Value = 0;
            //       prog.Visible = false;
        }

        public void EcrireDataA004(int NumSoc, string file, ProgressBar prog)
        {
            StreamReader sr = new StreamReader(file);
            string line = null;
            prog.Visible = true;
            prog.Minimum = 0;
            int ligne = 0;
            string matricule = string.Empty;
            string nom = string.Empty;
            string prenom = string.Empty;
            while ((line = sr.ReadLine()) != null)
            {
                ligne = ligne + 1;
                prog.Maximum = ligne;
                if (!(line.Equals("")))
                {
                    matricule = line.Substring(0, 10).Trim();
                    nom = line.Substring(10, 80).Trim().Replace("'", "''");
                    prenom = line.Substring(90, 20).Trim().Replace("'", "''");
                    if ((matricule != "") && (nom != "") && (prenom != ""))
                    {
                        string DateEmbauche = line.Substring(110, 8).Trim();
                        string DateDept = line.Substring(118, 8).Trim();
                        string CodeSce = line.Substring(126, 10).Trim();
                        string IntituleSce = line.Substring(136, 30).Trim().Replace("'", "''");
                        string NatureContrat = line.Substring(166, 10).Trim();
                        string Emploi = line.Substring(176, 30).Trim().Replace("'", "''");
                        string SalBase = line.Substring(206, 12).Trim();
                        string  CL01= line.Substring(218, 12).Trim();
                        string  A1100= line.Substring(230, 12).Trim();
                        string  A2200= line.Substring(242, 12).Trim();
                        string  A2100= line.Substring(254, 12).Trim();
                        string  A3411= line.Substring(266, 12).Trim();
                        string  A3410= line.Substring(278, 12).Trim();
                        string  A3412= line.Substring(290, 12).Trim();
                        string  A3317= line.Substring(302, 12).Trim();
                        string  A3712= line.Substring(314, 12).Trim();
                        string  A3710= line.Substring(326, 12).Trim();
                        string  A3711= line.Substring(338, 12).Trim();
                        string  A3713= line.Substring(350, 12).Trim();
                        string  A4112= line.Substring(362, 12).Trim();
                        string  A4113= line.Substring(374, 12).Trim();
                        string  A4120= line.Substring(386, 12).Trim();

                        string A4216 = line.Substring(398, 12).Trim();

                        string A4222 = line.Substring(410, 12).Trim();

                        string A4361 = line.Substring(422, 12).Trim();

                        string A4215 = line.Substring(434, 12).Trim();

                        string A4320 = line.Substring(446, 12).Trim();

                        string A4370 = line.Substring(458, 12).Trim();

                        string A4430 = line.Substring(470, 12).Trim();

                        string A7113 = line.Substring(482, 12).Trim();

                        string BRUT = line.Substring(494, 12).Trim();

                        string AjusCot = line.Substring(506, 12).Trim();

                        string A9112 = line.Substring(518, 12).Trim();

                        string A6110 = line.Substring(530, 12).Trim();

                        string A6210 = line.Substring(542, 12).Trim();

                        string A9310 = line.Substring(554, 12).Trim();

                        string A6310 = line.Substring(566, 12).Trim();

                        string A6311 = line.Substring(578, 12).Trim();

                        string A6340 = line.Substring(590, 12).Trim();

                        string A6210A = line.Substring(602, 12).Trim();

                        string A6410 = line.Substring(614, 12).Trim();

                        string A6430 = line.Substring(626, 12).Trim();

                        string A6520 = line.Substring(638, 12).Trim();

                        string A6510 = line.Substring(650, 12).Trim();

                        string A6530 = line.Substring(662, 12).Trim();

                        string A6830 = line.Substring(674, 12).Trim();

                        string A6490 = line.Substring(686, 12).Trim();

                        string RUB_DEDSA = line.Substring(698, 12).Trim();

                        string NETPAIE = line.Substring(710, 12).Trim();

                        string A6110A = line.Substring(722, 12).Trim();

                        string A6120 = line.Substring(734, 12).Trim();

                        string A6150 = line.Substring(746, 12).Trim();

                        string A6210AA = line.Substring(758,12).Trim();

                        string RUB_COTEPY = line.Substring(770, 12).Trim();

                        string MoisPe = line.Substring(782, 9).Trim();

                        string MoisEnt = line.Substring(791, 2).Trim();

                        string A4811 = line.Substring(793, 12).Trim();

                        string A4171 = line.Substring(805, 12).Trim();

                        string A4380 = line.Substring(817, 12).Trim();

                        string A4812 = line.Substring(829 , 12).Trim();

                       

                        prog.Value = prog.Value + 1;
                        int Annee = 2015;//DateTime.Today.Year;
                        int Mois = int.Parse(MoisEnt);
                        service.insertGrossToNetA004(matricule, nom.Replace("�", "e"), prenom.Replace("�", "e"), DateEmbauche,
                            DateDept, CodeSce, IntituleSce.Replace("�", "e"), NatureContrat, Emploi.Replace("�", "e"), SalBase,
                            CL01, A1100, A2200, A2100, A3411, A3410, A3412, A3317, A3712, A3710, A3711, A3713, A4112, A4113, A4120, A4216,A4222, A4361,
                            A4215,A4320,A4370,A4430,A7113,BRUT,AjusCot,A9112,A6110,A6210,A9310,A6310,A6311,A6340,A6210A,A6410,
                            A6430, A6520, A6510, A6530, A6830,A6490, RUB_DEDSA, NETPAIE, A6110A, A6120, A6150, A6210AA, RUB_COTEPY,
                              MoisPe.Replace("�", "e"), NumSoc.ToString(), Mois, Annee,A4811,A4171,A4380,A4812);
                    }
                }

            }
            MessageBox.Show("Import realized with success");
            //       prog.Value = 0;
            //       prog.Visible = false;
        }

        public void EcrireDataA021(int NumSoc, string file, ProgressBar prog)
        {
            StreamReader sr = new StreamReader(file);
            string line = null;
            prog.Visible = true;
            prog.Minimum = 0;
            int ligne = 0;
            string[] split;
            string matricule = string.Empty;
            string nom = string.Empty;
            string prenom = string.Empty;
            int index ;

         
            while ((line = sr.ReadLine()) != null)
            {
               index=0;
                ligne = ligne + 1;
                prog.Maximum = ligne;
                if (!(line.Equals("")))
                {
                    split = line.Split(new[] { "   " }, StringSplitOptions.RemoveEmptyEntries);
                   

                    matricule =  split[index++].Trim();
                    nom =  split[index++].Trim().Replace("'", "''");
                    prenom = split[index++].Trim().Replace("'", "''");
                    if ((matricule != "") && (nom != "") && (prenom != ""))
                    {
                      
                        string Datedembauchesociete = split[index++].Trim();
                        string Datededepartsociete = "";
                        string Intituledepartement = split[index++].Replace("'", "''"); ;
                        string CodeCostcenter = split[index++].Trim();
                        string Intitulecostcenter = split[index++].Trim();
                        string Intituledelanatureducontrat = split[index++].Trim();
                        string Emploioccupe = split[index++].Replace("'", "''"); ;
                        string Salairedebase = split[index++].Trim();
                        string Nbjoursdepresence = split[index++].Trim();
                        string SalairedebaseJtrvaille = split[index++].Trim();
                        string Inddepresence = split[index++].Trim();
                        string InddeTransport = split[index++].Trim();
                        string PrimeComplementairedePresence = split[index++].Trim();
                        string IndemnitedeVoiture = split[index++].Trim();
                        string IndFixederepresentation = split[index++].Trim();
                        string Indemnitedexpatriation = split[index++].Trim();
                      string  HS125HS0= split[index++].Trim();
                      string  ValeurHS0= split[index++].Trim();
                      string  HS1502= split[index++].Trim();
                       string ValeurHS02= split[index++].Trim();
                       string HS175HS06	= split[index++].Trim();
                      string  ValeurHS06= split[index++].Trim();
                      string  HS200HS03= split[index++].Trim();
                       string ValeurHS03= split[index++].Trim();
                       string HsupplementairenuitHS04= split[index++].Trim();
                      string  valeurHsuplementairenuitHS04= split[index++].Trim();
                      string  primeLVP= split[index++].Trim();
                      string  PrimeSIP= split[index++].Trim();
                      string  PrimeAOS= split[index++].Trim();
                       string Bonus= split[index++].Trim();
                     string   Compensationsupport= split[index++].Trim();
                      string  Rappelsursalaire= split[index++].Trim();
                      string  Rappelprimedetransport= split[index++].Trim();
                      string  PrimeperequationRC= split[index++].Trim();
                      string  PrimeperequationAV= split[index++].Trim();
                       string PrimeAid	= split[index++].Trim();
                      string  Primemariage= split[index++].Trim();
                      string  Primedenaissance= split[index++].Trim();
                      string  Congespayes= split[index++].Trim();
                      string  Indemnitedepreavis= split[index++].Trim();
                      string  Gratificationfindeservice= split[index++].Trim();
                      string  Indemnitedelicenciement= split[index++].Trim();
                      string  Avticketsrestaurant= split[index++].Trim();
                      string  AvAssurancemaladie= split[index++].Trim();
                      string  Avassuranceretraitecomplementaire= split[index++].Trim();
                      string  Avcarburant= split[index++].Trim();
                      string  Avvoiture= split[index++].Trim();
                      string  Avscolariteenfants= split[index++].Trim();
                     string   Avutiliteexpat= split[index++].Trim();
                     string   Avlogement= split[index++].Trim();
                       string SalaireBrut= split[index++].Trim();
                       string BrutEricsson= split[index++].Trim();
                      string  CNSSEmploye= split[index++].Trim();
                      string  AssuranceretraitecomplemCNRPSemploye= split[index++].Trim();
                      string  AssuranceretraiteComplementaire= split[index++].Trim();
                     string   Salaireimposable= split[index++].Trim();
                      string  IRPP= split[index++].Trim();
                      string  Redevance1= split[index++].Trim();
                      string  PrêtCNSS= split[index++].Trim();
                      string  AutresRetenuesursalaire= split[index++].Trim();
                      string  DeductionAvTR= split[index++].Trim();
                      string  DeductionAvassurancemaladie	= split[index++].Trim();
                      string  DeductionAvassurancecomplementaire	= split[index++].Trim();
                      string  DeductionAvVoiture= split[index++].Trim();
                      string  DeductionAvCarburant= split[index++].Trim();
                      string  DeductionAvloyer= split[index++].Trim();
                       string DeductionAvIndExpat	= split[index++].Trim();
                       string Deductionutiliteexpat= split[index++].Trim();
                      string  NetàPayer= split[index++].Trim();
                      string  CNSSEmployeur= split[index++].Trim();
                      string  Accidentdetravail= split[index++].Trim();
                      string  AssurancemaladieEmployeur= split[index++].Trim();
                      string  AssuranceComplementaireEmployeur= split[index++].Trim();
                      string  AssuranceretraitcompleCNRPSemplye= split[index++].Trim();
                      string  TFP= split[index++].Trim();
                      string  FOPROLOS= split[index++].Trim();
                       string TotalContributionEmployeur= split[index++].Substring(0,8).Trim();
                     //  string Moisdepaie = "Janvier";
                       string Moisdepaie = split[index-1].Replace(TotalContributionEmployeur, "").Trim();

                        prog.Value = prog.Value + 1;
                        FormatDate f = new FormatDate();
                        int Annee = DateTime.Today.Year;
                        int Mois = f.GetMonthByName(Moisdepaie);//int.Parse(MoisEnt);
                        service.insertGrossToNetA021(matricule,
             nom.Replace("�", "e"),
             prenom.Replace("�", "e"),
                 Datedembauchesociete,
                         Datededepartsociete,
                         Intituledepartement,
                         CodeCostcenter,
                         Intitulecostcenter.Replace("�", "e"),
                         Intituledelanatureducontrat,
                         Emploioccupe.Replace("�", "e"),
                         Salairedebase,
                         Nbjoursdepresence,
                         SalairedebaseJtrvaille,
                         Inddepresence,
                         InddeTransport,
                         PrimeComplementairedePresence,
                         IndemnitedeVoiture,
                         IndFixederepresentation,
                         Indemnitedexpatriation,
                        HS125HS0,
                        ValeurHS0,
                        HS1502,
                        ValeurHS02,
                        HS175HS06,
                        ValeurHS06,
                        HS200HS03,
                        ValeurHS03,
                        HsupplementairenuitHS04,
                        valeurHsuplementairenuitHS04,
                        primeLVP,
                        PrimeSIP,
                        PrimeAOS,
                        Bonus,
                        Compensationsupport,
                        Rappelsursalaire,
                        Rappelprimedetransport,
                        PrimeperequationRC,
                        PrimeperequationAV,
                        PrimeAid,
                        Primemariage,
                        Primedenaissance,
                        Congespayes,
                        Indemnitedepreavis,
                        Gratificationfindeservice,
                        Indemnitedelicenciement,
                        Avticketsrestaurant,
                        AvAssurancemaladie,
                        Avassuranceretraitecomplementaire,
                        Avcarburant,
                        Avvoiture,
                        Avscolariteenfants,
                        Avutiliteexpat,
                        Avlogement,
                        SalaireBrut,
                        BrutEricsson,
                        CNSSEmploye,
                        AssuranceretraitecomplemCNRPSemploye,
                        AssuranceretraiteComplementaire,
                        Salaireimposable,
                        IRPP,
                        Redevance1,
                        PrêtCNSS,
                        AutresRetenuesursalaire,
                        DeductionAvTR,
                        DeductionAvassurancemaladie,
                        DeductionAvassurancecomplementaire,
                        DeductionAvVoiture,
                        DeductionAvCarburant,
                        DeductionAvloyer,
                        DeductionAvIndExpat,
                        Deductionutiliteexpat,
                        NetàPayer,
                        CNSSEmployeur,
                        Accidentdetravail,
                        AssurancemaladieEmployeur,
                        AssuranceComplementaireEmployeur,
                        AssuranceretraitcompleCNRPSemplye,
                        TFP,
                        FOPROLOS,
                        TotalContributionEmployeur,
                        Moisdepaie.Replace("�", "e")
);
                    }
                }

            }
            MessageBox.Show("Import realized with success");
            //       prog.Value = 0;
            //       prog.Visible = false;
        }



        public void EcrireDataA010(int NumSoc, string file, ProgressBar prog)
        {
            StreamReader sr = new StreamReader(file);
            string line = null;
            prog.Visible = true;
            prog.Minimum = 0;
            int ligne = 0;
            string idOracle = string.Empty;
            string matricule = string.Empty;
            string nom = string.Empty;
            string prenom = string.Empty;
            while ((line = sr.ReadLine()) != null)
            {
                ligne = ligne + 1;
                prog.Maximum = ligne;
                if (!(line.Equals("")))
                {
                    idOracle = line.Substring(0, 10).Trim();
                    matricule = line.Substring(10, 10).Trim();
                    nom = line.Substring(20, 20).Trim().Replace("'", "''");
                    prenom = line.Substring(40, 20).Trim().Replace("'", "''");
                    if ((idOracle != "") && (nom != "") && (prenom != ""))
                    {
                        string DateEmbauche = line.Substring(60, 8).Trim();
                        string DateDept = line.Substring(68, 8).Trim();
                        string Emploi = line.Substring(76, 30).Trim().Replace("'", "''");
                        string NatureContrat = line.Substring(106, 10).Trim().Replace("'", "''");
                        string CodeDept = line.Substring(116, 30).Trim();
                        string Dept = line.Substring(146, 30).Trim().Replace("'", "''");
                        string HorBaseSal = line.Substring(176, 12).Trim();
                        string BullMod = line.Substring(188, 4).Trim();
                        string Categorie= line.Substring(192, 15).Trim();
                        string A1100 = line.Substring(207, 12).Trim();
                        string A1300 = line.Substring(219, 12).Trim();
                        string A2100 = line.Substring(231, 12).Trim();
                        string A2200 = line.Substring(243, 12).Trim();
                        string A3310 = line.Substring(255, 12).Trim();
                        string ABS_CONGNR = line.Substring(267, 12).Trim();
                        string A4174 = line.Substring(279, 12).Trim();
                        string RUB_ARR = line.Substring(291, 12).Trim();
                        string A4373= line.Substring(303, 12).Trim();
                        string HS01 = line.Substring(315, 12).Trim();
                        string A4110 = line.Substring(327, 12).Trim();
                        string HS02 = line.Substring(339, 12).Trim();
                        string A4111 = line.Substring(351, 12).Trim();
                        string HS03 = line.Substring(363, 12).Trim();
                        string A4113 = line.Substring(375, 12).Trim();
                        string HS06 = line.Substring(387, 12).Trim();
                        string A4112 = line.Substring(399, 12).Trim();
                        string HS04 = line.Substring(411, 12).Trim();
                        string A4120 = line.Substring(423, 12).Trim();
                        string HS07 = line.Substring(435, 12).Trim();
                        string A4372 = line.Substring(447, 12).Trim();
                        string HS05 = line.Substring(459, 12).Trim();
                        string A4114 = line.Substring(471, 12).Trim();
                        string HA04 = line.Substring(483, 12).Trim();
                        string HA06 = line.Substring(495, 12).Trim();
                        string RUB_JCHOM = line.Substring(507, 12).Trim();
                        string HA07 = line.Substring(519, 12).Trim();
                        string RUB_TRP = line.Substring(531, 12).Trim();
                        string HA09 = line.Substring(543, 12).Trim();
                        string A5500 = line.Substring(555, 12).Trim();
                        string A4222 = line.Substring(567, 12).Trim();
                        string A4371 = line.Substring(579, 12).Trim();
                        string A4130 = line.Substring(591, 12).Trim();
                        string A4135 = line.Substring(603, 12).Trim();
                        string A2361 = line.Substring(615, 12).Trim();
                        string A4816 = line.Substring(627, 12).Trim();
                        string A4810 = line.Substring(639, 12).Trim();
                        string HA01 = line.Substring(651, 12).Trim();
                        string  A4811= line.Substring(663, 12).Trim();
                        string A4812 = line.Substring(675, 12).Trim();
                        string A5100= line.Substring(687, 12).Trim();
                        string A5400= line.Substring(699, 12).Trim();
                        string A5900 = line.Substring(711, 12).Trim();
                        string A5200 = line.Substring(723, 12).Trim();
                        string A6510 = line.Substring(735, 12).Trim();
                        string A3309 = line.Substring(747, 12).Trim();
                        string A4817 = line.Substring(759, 12).Trim();
                        string A4818 = line.Substring(771, 12).Trim();
                        string A4313 = line.Substring(783, 12).Trim();
                        string A4312 = line.Substring(795, 12).Trim();
                        string TESTSTC = line.Substring(807, 12).Trim();
                        string A4800 = line.Substring(819, 12).Trim();
                        string PRES8 = line.Substring(831, 12).Trim();
                        string PRES2 = line.Substring(843, 12).Trim();
                        string A7115= line.Substring(855, 12).Trim();
                        string BRUT = line.Substring(867, 12).Trim();
                        string A6110 = line.Substring(879, 12).Trim();
                        string A9310 = line.Substring(891, 12).Trim();
                        string A6310 = line.Substring(903, 12).Trim();
                        string A6311 = line.Substring(915, 12).Trim();
                        string A6490 = line.Substring(927, 12).Trim();
                        string A6491 = line.Substring(939, 12).Trim();
                        string A6492 = line.Substring(951, 12).Trim();
                        string A6840 = line.Substring(963, 12).Trim();
                        string A6610 = line.Substring(975, 12).Trim();
                        string A6611 = line.Substring(987, 12).Trim();
                        string A6613 = line.Substring(999, 12).Trim();
                        string A6612 = line.Substring(1011, 12).Trim();
                        string A6410 = line.Substring(1023, 12).Trim();
                        string A6511 = line.Substring(1035, 12).Trim();
                        string A6614 = line.Substring(1047, 12).Trim();
                        string RUB_DEDSA = line.Substring(1059, 12).Trim();
                        string NETPAIE = line.Substring(1071, 12).Trim();
                        string A6110A = line.Substring(1083, 12).Trim();
                        string A6150 = line.Substring(1095, 12).Trim();
                        string A7115A = line.Substring(1107, 12).Trim();
                        string RUB_COTEPY = line.Substring(1119, 12).Trim();
                        string ModPe = line.Substring(1131, 10).Trim();
                        string MoisPe = line.Substring(1141, 9).Trim();
                        string MoisEnt = line.Substring(1150, 2).Trim();
                        string A4175 = line.Substring(1152, 12).Trim();
                        string A4374 = line.Substring(1164, 12).Trim();
                        prog.Value = prog.Value + 1;
                        int Annee = DateTime.Today.Year;
                        int Mois = int.Parse(MoisEnt);
                        service.insertGrossToNetA010(idOracle,matricule, nom.Replace("�", "e"), prenom.Replace("�", "e"), DateEmbauche,
                            DateDept, Emploi.Replace("�", "e"), NatureContrat.Replace("�", "e"), CodeDept, Dept.Replace("�", "e"), HorBaseSal,
                            BullMod, Categorie, A1100, A1300, A2100, A2200, A3310, ABS_CONGNR, A4174, RUB_ARR, A4373, HS01, A4110, HS02, A4111, HS03, A4113,
                            HS06,A4112,HS04,A4120,HS07,A4372,HS05,A4114,HA04,HA06,RUB_JCHOM,HA07,RUB_TRP,HA09,A5500,A4222,A4371,A4130,A4135,A2361,
                            A4816,A4810,HA01,A4811,A4812,A5100,A5400,A5900,A5200,A6510,A3309,A4817,A4818,A4313,A4312,TESTSTC,A4800,PRES8,PRES2,
                            A7115,BRUT,A6110,A9310,A6310,A6311,A6490,A6491,A6492,A6840,A6610,A6611,A6613,A6612,A6410,A6511,A6614,RUB_DEDSA,NETPAIE,
                            A6110A,A6150,A7115A,RUB_COTEPY,ModPe,MoisPe.Replace("�", "e"), NumSoc, Mois, Annee,A4175,A4374);
                    }
                }

            }
            MessageBox.Show("Import realized with success");
            //       prog.Value = 0;
            //       prog.Visible = false;
        }

        public void EcrireDataA011(int NumSoc, string file, ProgressBar prog)
        {
            StreamReader sr = new StreamReader(file);
            string line = null;
            prog.Visible = true;
            prog.Minimum = 0;
            int ligne = 0;
            string matricule = string.Empty;
            string nom = string.Empty;
            string prenom = string.Empty;
            while ((line = sr.ReadLine()) != null)
            {
                ligne = ligne + 1;
                prog.Maximum = ligne;
                if (!(line.Equals("")))
                {
                    matricule = line.Substring(0, 10).Trim();
                    nom = line.Substring(10, 20).Trim().Replace("'", "''");
                    prenom = line.Substring(30, 20).Trim().Replace("'", "''");
                    if ((matricule != "") && (nom != "") && (prenom != ""))
                    {
                        string DateEmbauche = line.Substring(50, 8).Trim();
                        string DateDepart = line.Substring(58, 8).Trim();
                        string IntituleDept = line.Substring(66, 10).Trim().Replace("'", "''");
                        string IntituleSce = line.Substring(76, 10).Trim().Replace("'", "''");
                        string IntituleCat = line.Substring(86, 10).Trim().Replace("'", "''");
                        string NatContrat = line.Substring(96, 10).Trim();
                        string Emploi = line.Substring(106, 30).Trim().Replace("'", "''");
                        string SalBase = line.Substring(136, 12).Trim();
                        string CL01 = line.Substring(148, 12).Trim();
                        string A1200 = line.Substring(160, 12).Trim();
                        string A2100 = line.Substring(172, 12).Trim();
                        string A2200 = line.Substring(184, 12).Trim();
                        string A2311 = line.Substring(196, 12).Trim();
                        string A3111 = line.Substring(208, 12).Trim();
                        string A3413 = line.Substring(220, 12).Trim();
                        string A3410 = line.Substring(232, 12).Trim();
                        string A4351 = line.Substring(244, 12).Trim();
                        string A4220 = line.Substring(256, 12).Trim();
                        string A4211 = line.Substring(268, 12).Trim();
                        string A4369 = line.Substring(280, 12).Trim();
                        string A4360 = line.Substring(292, 12).Trim();
                        string A4380 = line.Substring(304, 12).Trim();
                        string A4717 = line.Substring(316, 12).Trim();
                        string A4716 = line.Substring(328, 12).Trim();
                        string A4110 = line.Substring(340, 12).Trim();
                        string A4111 = line.Substring(353, 12).Trim();
                        string A4113 =  line.Substring(364, 12).Trim();
                        string A4114 =line.Substring(376, 12).Trim();
                        string A4812 = line.Substring(388, 12).Trim();
                        string A4171 = line.Substring(400, 12).Trim();
                        string BRUT = line.Substring(412, 12).Trim();
                        string A6110 = line.Substring(424, 12).Trim();
                        string A9121 = line.Substring(436, 12).Trim();
                        string A6310 = line.Substring(448, 12).Trim();
                        string A6311 = line.Substring(460, 12).Trim();
                        string TotDedEmp = line.Substring(472, 12).Trim();
                        string NETPAIE = line.Substring(484, 12).Trim();
                        string A6110A = line.Substring(496, 12).Trim();
                        string A6150 = line.Substring(508, 12).Trim();
                        string A6320 = line.Substring(520, 12).Trim();
                        string A6330 = line.Substring(532, 12).Trim();
                        string RUB_EPYC = line.Substring(544, 12).Trim();
                        string TotCostLab = line.Substring(556, 12).Trim();
                        string MoisPe = line.Substring(568, 9).Trim();
                        string MoisEnt = line.Substring(577, 2).Trim();
                        string CodeDept = line.Substring(579, 10).Trim();
                        
                        prog.Value = prog.Value + 1;
                        int Annee = DateTime.Today.Year;
                        int Mois = int.Parse(MoisEnt);
                        service.insertGrossToNetA011(matricule, nom.Replace("�", "e"), prenom.Replace("�", "e"), DateEmbauche,
                        DateDepart, IntituleDept.Replace("�", "e"), IntituleSce.Replace("�", "e"), IntituleCat.Replace("�", "e"),
                        NatContrat, Emploi.Replace("�", "e"),SalBase,CL01,A1200,A2100,A2200,A2311,A3111,A3413,A3410,A4351,A4220,A4211,
                        A4369, A4360, A4380, A4717, A4716, A4110, A4111, A4113, A4114,A4812,A4171, BRUT, A6110, A9121, A6310, A6311, TotDedEmp, NETPAIE,
                        A6110A,A6150,A6320,A6330,RUB_EPYC,TotCostLab,MoisPe.Replace("�", "e"), NumSoc.ToString(), Mois, Annee,CodeDept);
                    }
                }

            }
            MessageBox.Show("Import realized with success");
            //       prog.Value = 0;
            //       prog.Visible = false;
        }

        public void EcrireDataA014(int NumSoc, string file, ProgressBar prog)
        {
            StreamReader sr = new StreamReader(file);
            string line = null;
            prog.Visible = true;
            prog.Minimum = 0;
            int ligne = 0;
            string idOracle = string.Empty;
            string matricule = string.Empty;
            string nom = string.Empty;
            string prenom = string.Empty;
            while ((line = sr.ReadLine()) != null)
            {
                ligne = ligne + 1;
                prog.Maximum = ligne;
                if (!(line.Equals("")))
                {
                    idOracle = line.Substring(0, 10).Trim();
                    matricule = line.Substring(10, 10).Trim();
                    nom = line.Substring(20, 20).Trim().Replace("'", "''");
                    prenom = line.Substring(40, 20).Trim().Replace("'", "''");
                    if ((idOracle != "") && (nom != "") && (prenom != ""))
                    {
                        string DateEmbauche = line.Substring(60, 8).Trim();
                        string DateDept = line.Substring(68, 8).Trim();
                        string ExpatLoc = line.Substring(76, 10).Trim();
                        string Emploi = line.Substring(86, 30).Trim().Replace("'", "''");
                        string IntituleDept = line.Substring(116, 30).Trim().Replace("'", "''");
                        string CentreCout = line.Substring(146, 13).Trim();
                        string A1000 = line.Substring(159, 12).Trim();
                        string A1910 = line.Substring(171, 12).Trim();
                        string A1920 = line.Substring(183, 12).Trim();
                        string A3310 = line.Substring(195, 12).Trim();
                        string A3317 = line.Substring(207, 12).Trim();
                        string A3711 = line.Substring(219, 12).Trim();
                        string HS01 = line.Substring(231, 12).Trim();
                        string A4110 = line.Substring(243, 12).Trim();
                        string HS02 = line.Substring(255, 12).Trim();
                        string A4111 = line.Substring(267, 12).Trim();
                        string HS03 = line.Substring(279, 12).Trim();
                        string A4113 = line.Substring(291, 12).Trim();
                        string A4115 = line.Substring(304, 12).Trim();
                        string A4100 = line.Substring(315, 12).Trim();
                        string A4170 = line.Substring(327, 12).Trim();
                        string A4159 = line.Substring(339, 12).Trim();
                        string A4175 = line.Substring(351, 12).Trim();
                        string A4224 = line.Substring(364, 12).Trim();
                        string A4411 = line.Substring(375, 12).Trim();
                        string A4816 = line.Substring(387, 12).Trim();
                        string A4817 = line.Substring(399, 12).Trim();
                        string A4222 = line.Substring(411, 12).Trim();
                        string A4340 = line.Substring(423, 12).Trim();
                        string A4360 = line.Substring(435, 12).Trim();
                        string A4361 = line.Substring(448, 12).Trim();
                        string A4380 = line.Substring(459, 12).Trim();
                        string A4381 = line.Substring(471, 12).Trim();
                        string A4410 = line.Substring(483, 12).Trim();
                        string A4412 = line.Substring(495, 12).Trim();
                        string A4450 = line.Substring(507, 12).Trim();
                        string A4460 = line.Substring(519, 12).Trim();
                        string A4550 = line.Substring(531, 12).Trim();
                        string A4711 = line.Substring(543, 12).Trim();
                        string A4712 = line.Substring(555, 12).Trim();
                        string A4715 = line.Substring(557, 12).Trim();
                        string A4716 = line.Substring(579, 12).Trim();
                        string A4717 = line.Substring(591, 12).Trim();
                        string A4790 = line.Substring(603, 12).Trim();
                        string A4811 = line.Substring(615, 12).Trim();
                        string A4812 = line.Substring(627, 12).Trim();
                        string A4813 = line.Substring(639, 12).Trim();
                        string IFRT = line.Substring(651, 12).Trim();
                        string IndemRet = line.Substring(663, 12).Trim();
                        string Rappel = line.Substring(675, 12).Trim();
                        string A5420 = line.Substring(687, 12).Trim();
                        string A5200 = line.Substring(699, 12).Trim();
                        string PrimeTrim = line.Substring(711, 12).Trim();
                        string CL03 = line.Substring(723, 12).Trim();
                        string A4171 = line.Substring(735, 12).Trim();
                        string A7193 = line.Substring(747, 12).Trim();
                        string A7114 = line.Substring(759, 12).Trim();
                        string A7117 = line.Substring(771, 12).Trim();
                        string A7191 = line.Substring(783, 12).Trim();
                        string A7192 = line.Substring(795, 12).Trim();
                        string A7197 = line.Substring(807, 12).Trim();
                        string A7196 = line.Substring(820, 12).Trim();
                        string A7198 = line.Substring(831, 12).Trim();
                        string A7199 = line.Substring(843, 12).Trim();
                        string A7692 = line.Substring(855, 12).Trim();
                        string A7115 = line.Substring(867, 12).Trim();
                        string A7118 = line.Substring(879, 12).Trim();
                        string BRUT = line.Substring(891, 12).Trim();
                        string A9112 = line.Substring(903, 12).Trim();
                        string A6110 = line.Substring(915, 12).Trim();
                        string A6120 = line.Substring(927, 12).Trim();
                        string A6201 = line.Substring(939, 12).Trim();
                        string A6221 = line.Substring(951, 12).Trim();
                        string A6222 = line.Substring(963, 12).Trim();                     
                        string A9121 = line.Substring(975, 12).Trim();
                        string A6310 = line.Substring(987, 12).Trim();
                        string A6311 = line.Substring(999, 12).Trim();
                        string A7115A = line.Substring(1011, 12).Trim();
                        string A7118A = line.Substring(1023, 12).Trim();
                        string A7193A = line.Substring(1035, 12).Trim();
                        string A7114A = line.Substring(1047, 12).Trim();
                        string A6837 = line.Substring(1059, 12).Trim();
                        string A6832 = line.Substring(1071, 12).Trim();
                        string A7197A = line.Substring(1083, 12).Trim();
                        string A6842 = line.Substring(1095, 12).Trim();
                        string A6843 = line.Substring(1107, 12).Trim();
                        string A6844 = line.Substring(1119, 12).Trim();
                        string A6850 = line.Substring(1131, 12).Trim();
                        string A6410 = line.Substring(1143, 12).Trim();
                        string A6430 = line.Substring(1155, 12).Trim();
                        string A6470 = line.Substring(1167, 12).Trim();
                        string A6491 = line.Substring(1179, 12).Trim();
                        string A6614 = line.Substring(1191, 12).Trim();
                        string A6530 = line.Substring(1203, 12).Trim();
                        string A6510 = line.Substring(1215, 12).Trim();
                        string A6521 = line.Substring(1227, 12).Trim();
                        string A6520 = line.Substring(1239, 12).Trim();
                        string A8201 = line.Substring(1251, 12).Trim();
                        string A8215 = line.Substring(1263, 12).Trim();
                        string A8217 = line.Substring(1275, 12).Trim();
                        string RUB_DEDSA = line.Substring(1287, 12).Trim();
                        string NETPAIE = line.Substring(1299, 12).Trim();
                        string A6110A = line.Substring(1311, 12).Trim();
                        string A6120A = line.Substring(1323, 12).Trim();
                        string A6150 = line.Substring(1335, 12).Trim();
                        string A7115AA = line.Substring(1347, 12).Trim();
                        string A7118AA = line.Substring(1359, 12).Trim();
                        string RUB_COTEPY = line.Substring(1371, 12).Trim();
                        string MoisPe = line.Substring(1383, 9).Trim();
                        string MoisEnt = line.Substring(1392, 2).Trim();
                        string A5400 = line.Substring(1394, 12).Trim();
                        string A4610 = line.Substring(1406, 12).Trim();

                        prog.Value = prog.Value + 1;
                        int Annee = DateTime.Today.Year;
                        int Mois = int.Parse(MoisEnt);
                        service.insertGrossToNetA014(idOracle, matricule, nom.Replace("�", "e"), prenom.Replace("�", "e"),
                        DateEmbauche, DateDept, ExpatLoc, Emploi.Replace("�", "e"),
          IntituleDept.Replace("�", "e"), CentreCout, A1000, A1910, A1920, A3310, A3317, A3711, HS01, A4110, HS02, A4111,
           HS03, A4113, A4115, A4100, A4170, A4159, A4175, A4224, A4411, A4816,A4817, A4222, A4340, A4360,
           A4361, A4380, A4381, A4410, A4412, A4450, A4460, A4550, A4711, A4712, A4715, A4716, A4717,
           A4790, A4811, A4812, A4813, IFRT, IndemRet, Rappel, A5420,A5200, PrimeTrim, CL03, A4171, A7193,
           A7114, A7117, A7191, A7192, A7197, A7196, A7198, A7199, A7692, A7115, A7118, BRUT, A9112,
           A6110, A6120, A6201, A6221,A6222, A9121, A6310, A6311, A7115A, A7118A, A7193A, A7114A, A6837, A6832,
           A7197A, A6842, A6843, A6844, A6850, A6410, A6430, A6470, A6491, A6614, A6530, A6510, A6521,
           A6520, A8201, A8215, A8217, RUB_DEDSA, NETPAIE, A6110A, A6120A, A6150, A7115AA, A7118AA, RUB_COTEPY,
           MoisPe, NumSoc, Mois, Annee,A5400,A4610);
                    }
                }

            }
            MessageBox.Show("Import realized with success");
            //       prog.Value = 0;
            //       prog.Visible = false;
        }

        public void EcrireDataA015(int NumSoc, string file, ProgressBar prog)
        {
            StreamReader sr = new StreamReader(file);
            string line = null;
            prog.Visible = true;
            prog.Minimum = 0;
            int ligne = 0;
            string matricule = string.Empty;
            string nom = string.Empty;
            string prenom = string.Empty;
            while ((line = sr.ReadLine()) != null)
            {
                ligne = ligne + 1;
                prog.Maximum = ligne;
                if (!(line.Equals("")))
                {
                    matricule = line.Substring(0, 10).Trim();
                    nom = line.Substring(10, 20).Trim().Replace("'", "''");
                    prenom = line.Substring(30, 20).Trim().Replace("'", "''");
                    if ((matricule != "") && (nom != "") && (prenom != ""))
                    {
                        string DateEmbauche = line.Substring(50, 8).Trim();
                        string DateDepart = line.Substring(58, 8).Trim();
                        string CodeDept = line.Substring(66, 10).Trim().Replace("'", "''");
                        string IntituleDept = line.Substring(76, 32).Trim().Replace("'", "''");
                        string NatContrat = line.Substring(108, 10).Trim().Replace("'", "''");
                        string Emploi = line.Substring(118, 32).Trim().Replace("'", "''"); ;
                        string SalBaseAnn = line.Substring(150, 12).Trim();
                        string SALBASE_M = line.Substring(162, 12).Trim();
                        string CL01 = line.Substring(174, 12).Trim();
                        string A1200 = line.Substring(186, 12).Trim();
                        string A3410 = line.Substring(198, 12).Trim();
                        string A3110 = line.Substring(210, 12).Trim();
                        string A3711 = line.Substring(222, 12).Trim();
                        string A4222 = line.Substring(234, 12).Trim();
                        string A4362 = line.Substring(246, 12).Trim();
                        string A4311 = line.Substring(258, 12).Trim();
                        string A4380 = line.Substring(270, 12).Trim();
                        string A4711 = line.Substring(282, 12).Trim();
                        string A4717 = line.Substring(294, 12).Trim();
                        string A4718 = line.Substring(306, 12).Trim();
                        string A4364 = line.Substring(318, 12).Trim();
                        string A4171 = line.Substring(330, 12).Trim();

                        string A5100 = line.Substring(342, 12).Trim();

                        string A7111 = line.Substring(354, 12).Trim();

                        string A7112 = line.Substring(366, 12).Trim();

                        string A7118 = line.Substring(378, 12).Trim();

                        string BRUT = line.Substring(390, 12).Trim();

                        string A6110 = line.Substring(402, 12).Trim();

                        string A6120 = line.Substring(414, 12).Trim();

                        string A9121 = line.Substring(426, 12).Trim();

                        string A6310 = line.Substring(438, 12).Trim();

                        string A6311 = line.Substring(450, 12).Trim();
                        string A8000 = line.Substring(462, 12).Trim();
                        string A6850 = line.Substring(474, 12).Trim();
                        string A6220 = line.Substring(486, 12).Trim();
                        string RUB_DEDSA = line.Substring(498, 12).Trim();
                        string NETPAIE = line.Substring(510, 12).Trim();
                        string A6110A = line.Substring(522, 12).Trim();
                        string A6120A = line.Substring(534, 12).Trim();
                        string A6150 = line.Substring(546, 12).Trim();
                        string RUB_EPYC = line.Substring(558, 12).Trim();
                        string MoisPe = line.Substring(570, 9).Trim();
                        string MoisEnt = line.Substring(579, 2).Trim();

                        prog.Value = prog.Value + 1;
                        int Annee = DateTime.Today.Year;
                        int Mois = int.Parse(MoisEnt);
                        service.insertGrossToNetA015(matricule, nom.Replace("�", "e"), prenom.Replace("�", "e"),
                            DateEmbauche, DateDepart, CodeDept, IntituleDept, NatContrat, Emploi,
                        SalBaseAnn, SALBASE_M, CL01, A1200, A3410, A3110, A3711, A4222, A4362, A4311,
                      A4380, A4711, A4717, A4718, A4364, A4171,A5100, A7111, A7112, A7118, BRUT,
                        A6110, A6120, A9121, A6310, A6311, A8000,A6850, A6220, RUB_DEDSA, NETPAIE, A6110A, A6120A,
                       A6150, RUB_EPYC, MoisPe.Replace("�", "e"), NumSoc.ToString(), Mois, Annee);
           
                    }
                }

            }
            MessageBox.Show("Import realized with success");
            //       prog.Value = 0;
            //       prog.Visible = false;
        }

        public string VerifVide(string Vel)
        {
            string ValF = string.Empty;
            if (Vel == "")
            {
                ValF = string.Empty;
            }
            else
            {
                ValF = Vel;
            }
            return ValF;
        }

        public void EcrireDataA013(int NumSoc, string file, ProgressBar prog)
        {
            StreamReader sr = new StreamReader(file);
            string line = null;
            prog.Visible = true;
            prog.Minimum = 0;
            int ligne = 0;
            string idOracle = string.Empty;
            string matricule = string.Empty;
            string nom = string.Empty;
            string prenom = string.Empty;
            while ((line = sr.ReadLine()) != null)
            {
                ligne = ligne + 1;
                prog.Maximum = ligne;
                if (!(line.Equals("")))
                {
                    idOracle = line.Substring(0, 10).Trim();
                    matricule = VerifVide(line.Substring(10, 10).Trim());
                    nom = line.Substring(20, 20).Trim().Replace("'", "''");
                    prenom = line.Substring(40, 20).Trim().Replace("'", "''");
                    if ((idOracle != "") && (nom != "") && (prenom != ""))
                    {
                        string DateEmbauche = line.Substring(60, 8).Trim();
                        string DateDept = line.Substring(68, 8).Trim();
                        string Emploi = VerifVide(line.Substring(76, 32).Trim().Replace("'", "''"));
                        string NatureContrat = VerifVide(line.Substring(108, 10).Trim().Replace("'", "''"));
                        string CodeDept = line.Substring(118, 30).Trim();
                        string Dept = line.Substring(148, 30).Trim().Replace("'", "''");
                        string HorBaseSal = line.Substring(178, 12).Trim();
                        string BullMod = line.Substring(190, 4).Trim();
                        string Categorie = line.Substring(194, 15).Trim();
                        string ModPe = line.Substring(209, 12).Trim();
                        string A1100 = line.Substring(219, 12).Trim();
                        string A1300 = line.Substring(231, 12).Trim();
                        string A1310 = line.Substring(243, 12).Trim();
                        string A2100 = line.Substring(255, 12).Trim();
                        string A2200 = line.Substring(267, 12).Trim();
                        string A3310 = line.Substring(279, 12).Trim();
                        string A3110 = line.Substring(291, 12).Trim();
                        string ABS_CONGNR = line.Substring(303, 12).Trim();
                        string A4174 = line.Substring(315, 12).Trim();
                        string HS01 = line.Substring(327, 12).Trim();
                        string A4110 = line.Substring(339, 12).Trim();
                        string HS02 = line.Substring(351, 12).Trim();
                        string A4111 = line.Substring(363, 12).Trim();
                        string HS03 = line.Substring(375, 12).Trim();
                        string A4113 = line.Substring(387, 12).Trim();
                        string HS04 = line.Substring(399, 12).Trim();
                        string A4120 = line.Substring(411, 12).Trim();
                        string HS07 = line.Substring(423, 12).Trim();
                        string A4372 = line.Substring(435, 12).Trim();
                        string HS05 = line.Substring(447, 12).Trim();
                        string A4114 = line.Substring(459, 12).Trim();
                        string HA04 = line.Substring(471, 12).Trim();
                        string HA06 = line.Substring(483, 12).Trim();
                        string A4175 = line.Substring(495, 12).Trim();
                        string HA07 = line.Substring(507, 12).Trim();
                        string A4176 = line.Substring(519, 12).Trim();
                        string A4130 = line.Substring(531, 12).Trim();
                        string A4135 = line.Substring(543, 12).Trim();
                        string A2361 = line.Substring(555, 12).Trim();
                        string A4222 = line.Substring(567, 12).Trim();
                        string A3309 = line.Substring(579, 12).Trim();
                        string A4373 = line.Substring(591, 12).Trim();
                        string A4371 = line.Substring(603, 12).Trim();
                        string A4817 = line.Substring(615, 12).Trim();
                        string A4818 = line.Substring(627, 12).Trim();
                        string A4313 = line.Substring(639, 12).Trim();
                        string A4311 = line.Substring(651, 12).Trim();
                        string HA09 = line.Substring(663, 12).Trim();
                        string A5500 = line.Substring(675, 12).Trim();
                        string A5100 = line.Substring(688, 12).Trim();
                        string A5400 = line.Substring(699, 12).Trim();
                        string A5200 = line.Substring(711, 12).Trim();
                        string A5900 = line.Substring(723, 12).Trim();
                        string HA01 = line.Substring(735, 12).Trim();
                        string A4816 = line.Substring(747, 12).Trim();
                        string A4810 = line.Substring(759, 12).Trim();
                        string A4811 = line.Substring(771, 12).Trim();
                        string A4812 = line.Substring(783, 12).Trim();
                        string TESTSTC = line.Substring(795, 12).Trim();
                        string A4800 = line.Substring(807, 12).Trim();
                        string PRES8 = line.Substring(819, 12).Trim();
                        string A7113 = line.Substring(831, 12).Trim();
                        string CL23 = line.Substring(843, 12).Trim();
                        string A7116 = line.Substring(855, 12).Trim();
                        string A7115 = line.Substring(867, 12).Trim();
                        string BRUT = line.Substring(879, 12).Trim();
                        string A6110 = line.Substring(891, 12).Trim();
                        string A9310 = line.Substring(903, 12).Trim();
                        string A6310 = line.Substring(915, 12).Trim();
                        string A6311 = line.Substring(927, 12).Trim();
                        string A6614 = line.Substring(939, 12).Trim();
                        string A6410 = line.Substring(951, 12).Trim();
                        string A6490 = line.Substring(963, 12).Trim();
                        string A6491 = line.Substring(975, 12).Trim();
                        string A6492 = line.Substring(987, 12).Trim();
                        string A6511 = line.Substring(999, 12).Trim();
                        string A6430 = line.Substring(1011, 12).Trim();
                        string A6431 = line.Substring(1023, 12).Trim();
                        string A6840 = line.Substring(1035, 12).Trim();
                        string RUB_DEDSA = line.Substring(1047, 12).Trim();
                        string NETPAIE = line.Substring(1059, 12).Trim();
                        string A61109 = line.Substring(1071, 12).Trim();
                        string A6150 = line.Substring(1083, 12).Trim();
                        string A71159 = line.Substring(1095, 12).Trim();
                        string RUB_COTEPY = line.Substring(1107, 12).Trim();
                        string MoisPe = line.Substring(1119, 9).Trim();
                        string MoisEnt = line.Substring(1128, 2).Trim();
                        string A4374 = line.Substring(1130, 12).Trim();
                        prog.Value = prog.Value + 1;
                        int Annee = DateTime.Today.Year;
                        int Mois = int.Parse(MoisEnt);
                        service.insertGrossToNetA013( idOracle,  matricule,  nom.Replace("�", "e"),  prenom.Replace("�", "e"),  DateEmbauche,
                       DateDept, Emploi.Replace("�", "e"), NatureContrat, CodeDept, Dept.Replace("�", "e"), HorBaseSal, BullMod,
                       Categorie.Replace("�", "e"),ModPe, A1100, A1300, A1310, A2100, A2200, A3310, A3110, ABS_CONGNR, A4174, HS01, A4110, HS02, A4111, HS03, A4113, HS04, A4120,
                        HS07, A4372, HS05, A4114, HA04, HA06, A4175, HA07, A4176, A4130, A4135, A2361, A4222, A3309, A4373, A4371, A4817, A4818, A4313,
                        A4311, HA09, A5500, A5100, A5400, A5200, A5900, HA01, A4816, A4810, A4811, A4812, TESTSTC, A4800, PRES8, A7113, CL23, A7116, A7115,
                        BRUT, A6110, A9310, A6310, A6311, A6614, A6410, A6490, A6491, A6492, A6511, A6430, A6431, A6840, RUB_DEDSA, NETPAIE, A61109, A6150, A71159,
                        RUB_COTEPY, MoisPe, NumSoc, Mois, Annee,A4374);
                    }
                }

            }
            MessageBox.Show("Import realized with success");
            //       prog.Value = 0;
            //       prog.Visible = false;
        }

        public void EcrireDataA020(int NumSoc, string file, ProgressBar prog)
        {
            StreamReader sr = new StreamReader(file);
            string line = null;
            prog.Visible = true;
            prog.Minimum = 0;
            int ligne = 0;
            string idOracle = string.Empty;
            string matricule = string.Empty;
            string nom = string.Empty;
            string prenom = string.Empty;
            while ((line = sr.ReadLine()) != null)
            {
                ligne = ligne + 1;
                prog.Maximum = ligne;
                if (!(line.Equals("")))
                {
                    idOracle = line.Substring(0, 10).Trim();
                    matricule = line.Substring(10, 10).Trim();
                    nom = line.Substring(20, 20).Trim().Replace("'", "''");
                    prenom = line.Substring(40, 20).Trim().Replace("'", "''");
                    if ((matricule != "") && (nom != "") && (prenom != ""))
                    {
                        string DateEmbauche = line.Substring(60, 8).Trim();
                        string DateDept = line.Substring(68, 8).Trim();
                        string CodeDept = line.Substring(76, 10).Trim().Replace("'", "''");
                        string IntituleDept = line.Substring(86, 30).Trim().Replace("'", "''");
                        string ExpLoc = line.Substring(116, 10).Trim();
                        string Emploi = line.Substring(126, 30).Trim().Replace("'", "''");
                        string ContreCout = line.Substring(156, 13).Trim();                      
                        string A1000 = line.Substring(169, 12).Trim();
                        string A1910 = line.Substring(181, 12).Trim();
                        string A1920 = line.Substring(193, 12).Trim();
                        string A2311 = line.Substring(205, 12).Trim();
                        string A3317 = line.Substring(217, 12).Trim();
                        string A3510 = line.Substring(229, 12).Trim();
                        string A3711 = line.Substring(241, 12).Trim();
                        string HS01 = line.Substring(253, 12).Trim();
                        string A4110 = line.Substring(265, 12).Trim();
                        string HS02 = line.Substring(277, 12).Trim();
                        string A4111 = line.Substring(289, 12).Trim();
                        string HS03 = line.Substring(301, 12).Trim();
                        string A4113 = line.Substring(313, 12).Trim();
                        string A4115 = line.Substring(325, 12).Trim();
                        string A4100 = line.Substring(337, 12).Trim();
                        string A4159 = line.Substring(349, 12).Trim();
                        string A4175 = line.Substring(361, 12).Trim();
                        string A4224 = line.Substring(373, 12).Trim();
                        string A4160 = line.Substring(385, 12).Trim();
                        string A4222 = line.Substring(397, 12).Trim();
                        string A4340 = line.Substring(409, 12).Trim();
                        string A4360 = line.Substring(421, 12).Trim();
                        string A4361 = line.Substring(433, 12).Trim();
                        string A4380 = line.Substring(445, 12).Trim();
                        string A4381 = line.Substring(457, 12).Trim();
                        string A7117 = line.Substring(469, 12).Trim();
                        string A4410 = line.Substring(481, 12).Trim();
                        string A4816 = line.Substring(493, 12).Trim();
                        string A4817 = line.Substring(505, 12).Trim();
                        string A4412 = line.Substring(517, 12).Trim();
                        string A4450 = line.Substring(529, 12).Trim();
                        string A4460 = line.Substring(541, 12).Trim();
                        string A4550 = line.Substring(553, 12).Trim();

                        string A4711 = line.Substring(565, 12).Trim();

                        string A4712 = line.Substring(577, 12).Trim();

                        string A4715 = line.Substring(589, 12).Trim();

                        string A4716 = line.Substring(601, 12).Trim();

                        string A4717 = line.Substring(613, 12).Trim();

                        string A4790 = line.Substring(625, 12).Trim();

                        string A4811 = line.Substring(637, 12).Trim();

                        string A4812 = line.Substring(649, 12).Trim();

                        string A4813 = line.Substring(661, 12).Trim();

                        string IFRT = line.Substring(673, 12).Trim();

                        string IndemRetAnt = line.Substring(685, 12).Trim();

                        string LiquidIfrt = line.Substring(697, 12).Trim();

                        string A5100 = line.Substring(709, 12).Trim();

                        string A5420 = line.Substring(721, 12).Trim();

                        string A5200 = line.Substring(733, 12).Trim();

                        string PrimTrim = line.Substring(745, 12).Trim();

                        string COPRIMOI = line.Substring(757, 12).Trim();

                        string A4171 = line.Substring(769, 12).Trim();

                        string HA04 = line.Substring(781, 12).Trim();

                        string A4170 = line.Substring(793, 12).Trim();

                        string A7114 = line.Substring(805, 12).Trim();

                        string A7191 = line.Substring(817, 12).Trim();

                        string A7193 = line.Substring(829, 12).Trim();

                        string A7196 = line.Substring(841, 12).Trim();

                        string A7692 = line.Substring(853, 12).Trim();

                        string A7115 = line.Substring(865, 12).Trim();

                        string A7118 = line.Substring(877, 12).Trim();

                        string BRUT = line.Substring(889, 12).Trim();

                        string A9112 = line.Substring(901, 12).Trim();

                        string A6110 = line.Substring(913, 12).Trim();

                        string A6120 = line.Substring(925, 12).Trim();
                        string A6201 = line.Substring(937, 12).Trim();

                        string A6221 = line.Substring(949, 12).Trim();

                        string A9121 = line.Substring(961, 12).Trim();

                        string A6310 = line.Substring(973, 12).Trim();

                        string A6311 = line.Substring(985, 12).Trim();

                        string A6840 = line.Substring(997, 12).Trim();

                        string A6841 = line.Substring(1009, 12).Trim();

                        string A6835 = line.Substring(1021, 12).Trim();

                        string A6831 = line.Substring(1033, 12).Trim();

                        string A44509 = line.Substring(1045, 12).Trim();

                        string A6832 = line.Substring(1057, 12).Trim();

                        string A6834 = line.Substring(1069, 12).Trim();

                        string A6850 = line.Substring(1081, 12).Trim();

                        string A6410 = line.Substring(1093, 12).Trim();

                        string A6430 = line.Substring(1105, 12).Trim();

                        string A6470 = line.Substring(1117, 12).Trim();

                        string A6491 = line.Substring(1129, 12).Trim();

                        string A6520 = line.Substring(1141, 12).Trim();

                        string A6510 = line.Substring(1153, 12).Trim();

                        string A6521 = line.Substring(1165, 12).Trim();

                        string A6591 = line.Substring(1177, 12).Trim();

                        string A6614 = line.Substring(1189, 12).Trim();

                        string A8201 = line.Substring(1202, 12).Trim();

                        string A8215 = line.Substring(1213, 12).Trim();

                        string A8217 = line.Substring(1225, 12).Trim();

                        string RUB_DEDSA = line.Substring(1237, 12).Trim();

                        string NETPAIE = line.Substring(1249, 12).Trim();

                        string A61109 = line.Substring(1261, 12).Trim();

                        string A61209 = line.Substring(1273, 12).Trim();

                        string A6150 = line.Substring(1285, 12).Trim();

                        string A71159 = line.Substring(1297, 12).Trim();

                        string A62019 = line.Substring(1309, 12).Trim();

                        string RUB_COTEPY = line.Substring(1321, 12).Trim();

                        string MoisPaie = line.Substring(1333, 9).Trim();

                        string MoisEnt = line.Substring(1342, 2).Trim();

                        string A5400 = line.Substring(1344, 12).Trim();      
                      
                        prog.Value = prog.Value + 1;
                        int Annee = DateTime.Today.Year; 
                        int Mois = int.Parse(MoisEnt);
                        service.insertGrossToNetA020(idOracle,  matricule,  nom.Replace("�", "e"),  prenom.Replace("�", "e"),
                            DateEmbauche,  DateDept,  CodeDept,  IntituleDept.Replace("�", "e"),
                        ExpLoc,  Emploi.Replace("�", "e"),  ContreCout,  A1000,  A1910,  A1920,  A2311,  A3317,  A3510,  A3711,  HS01,  A4110,
                        HS02,  A4111,  HS03,  A4113,  A4115,  A4100,  A4159,  A4175,  A4224,  A4160,  A4222,  A4340,  A4360,
                        A4361, A4380, A4381, A7117, A4410, A4816,A4817, A4412, A4450, A4460, A4550, A4711, A4712,
                        A4715, A4716, A4717, A4790, A4811, A4812, A4813, IFRT, IndemRetAnt, LiquidIfrt, A5100,
                        A5420,A5200, PrimTrim, COPRIMOI, A4171, HA04, A4170, A7114, A7191, A7193, A7196, A7692, A7115,
                        A7118, BRUT, A9112, A6110, A6120, A6201, A6221, A9121, A6310, A6311, A6840, A6841, A6835,
                        A6831, A44509, A6832, A6834, A6850, A6410, A6430, A6470, A6491, A6520, A6510, A6521,
                        A6591, A6614, A8201, A8215, A8217, RUB_DEDSA, NETPAIE, A61109, A61209, A6150, A71159, A62019,
                        RUB_COTEPY, MoisPaie, NumSoc, Mois, Annee,A5400);
                    }
                }

            }
            MessageBox.Show("Import realized with success");
            //       prog.Value = 0;
            //       prog.Visible = false;
        }

        public void EcrireDataA016(int NumSoc, string file, ProgressBar prog)
        {
            StreamReader sr = new StreamReader(file);
            string line = null;
            prog.Visible = true;
            prog.Minimum = 0;
            int ligne = 0;
            string idOracle = string.Empty;
            string matricule = string.Empty;
            string nom = string.Empty;
            string prenom = string.Empty;
            while ((line = sr.ReadLine()) != null)
            {
                ligne = ligne + 1;
                prog.Maximum = ligne;
                if (!(line.Equals("")))
                {
                    matricule = line.Substring(0, 10).Trim();
                    nom = line.Substring(10, 80).Trim().Replace("'", "''");
                    prenom = line.Substring(90, 20).Trim().Replace("'", "''");
                    if ((matricule != "") && (nom != "") && (prenom != ""))
                    {
                        string DateEmbauche = line.Substring(110, 8).Trim();
                        string DateDept = line.Substring(118, 8).Trim();
                        string CodeDept = line.Substring(126, 10).Trim().Replace("'", "''");
                        string IntituleDept = line.Substring(136, 30).Trim().Replace("'", "''");
                        string NatureContrat = line.Substring(166, 60).Trim();
                        string Emploi = line.Substring(226, 30).Trim().Replace("'", "''");
                        string SALBASE = line.Substring(256, 12).Trim();
                        string HORAIRE = line.Substring(268, 12).Trim();
                        string A1000 = line.Substring(280, 12).Trim();
                        string A2100 = line.Substring(292, 12).Trim();
                        string A2200 = line.Substring(304, 12).Trim();
                        string A3413 = line.Substring(316, 12).Trim();
                        string A3414 = line.Substring(328, 12).Trim();
                        string A3415 = line.Substring(340, 12).Trim();
                        string A3418 = line.Substring(352, 12).Trim();
                        string A3316 = line.Substring(364, 12).Trim();
                        string A3317 = line.Substring(376, 12).Trim();
                        string A3510 = line.Substring(388, 12).Trim();
                        string A3520 = line.Substring(400, 12).Trim();
                        string A3530 = line.Substring(412, 12).Trim();
                        string A3417 = line.Substring(424, 12).Trim();
                        string A5910 = line.Substring(436, 12).Trim();
                        string A3318 = line.Substring(448, 12).Trim();
                        string A3610 = line.Substring(460, 12).Trim();
                        string A3620 = line.Substring(472, 12).Trim();
                        string A3630 = line.Substring(484, 12).Trim();
                        string A3640 = line.Substring(496, 12).Trim();
                        string A3651 = line.Substring(508, 12).Trim();
                        string A3652 = line.Substring(520, 12).Trim();
                        string A3650 = line.Substring(532, 12).Trim();
                        string A3660 = line.Substring(544, 12).Trim();
                        string A3661 = line.Substring(556, 12).Trim();
                        string A3670 = line.Substring(568, 12).Trim();
                        string A3680 = line.Substring(580, 12).Trim();
                        string A3681 = line.Substring(592, 12).Trim();
                        string A3682 = line.Substring(604, 12).Trim();
                        string A3683 = line.Substring(616, 12).Trim();
                        string A3684 = line.Substring(628, 12).Trim();
                        string A3685 = line.Substring(640, 12).Trim();
                        string BRUTFX = line.Substring(652, 12).Trim();
                        string A3686 = line.Substring(664, 12).Trim();
                        string A3690 = line.Substring(676, 12).Trim();
                        string CL31 = line.Substring(688, 12).Trim();
                        string A4460 = line.Substring(700, 12).Trim();
                        string CL32 = line.Substring(712, 12).Trim();
                        string A4461 = line.Substring(724, 12).Trim();
                        string CL33 = line.Substring(736, 12).Trim();
                        string A4462 = line.Substring(748, 12).Trim();
                        string CL34 = line.Substring(760, 12).Trim();
                        string A4463 = line.Substring(772, 12).Trim();
                        string CL35 = line.Substring(784, 12).Trim();
                        string A3110 = line.Substring(796, 12).Trim();
                        string CL37 = line.Substring(808, 12).Trim();
                        string A4292 = line.Substring(820, 12).Trim();
                        string A4293 = line.Substring(832, 12).Trim();
                        string CL39 = line.Substring(844, 12).Trim();
                        string A4294 = line.Substring(856, 12).Trim();
                        string HC01 = line.Substring(868, 12).Trim();
                        string A4295 = line.Substring(880, 12).Trim();
                        string CL40 = line.Substring(892, 12).Trim();
                        string A4296 = line.Substring(904, 12).Trim();
                        string A4491 = line.Substring(916, 12).Trim();
                        string A4162 = line.Substring(928, 12).Trim();
                        string A4391 = line.Substring(940, 12).Trim();
                        string A4381 = line.Substring(952, 12).Trim();
                        string A4840 = line.Substring(964, 12).Trim();
                        string A4382 = line.Substring(976, 12).Trim();
                        string A4850 = line.Substring(988, 12).Trim();
                        string A4383 = line.Substring(1000, 12).Trim();           
                        string A4860 = line.Substring(1012, 12).Trim();
                        string A4718 = line.Substring(1024, 12).Trim();     
                        string A4717 = line.Substring(1036, 12).Trim();
                        string A4713 = line.Substring(1048, 12).Trim();
                        string A4716 = line.Substring(1060, 12).Trim();
                        string HS01 = line.Substring(1072, 12).Trim();      
                        string A4111 = line.Substring(1084, 12).Trim();
                        string HS02 = line.Substring(1096, 12).Trim();
                        string A4112 = line.Substring(1108, 12).Trim();
                        string HS03 = line.Substring(1120, 12).Trim();
                        string A4113 = line.Substring(1132, 12).Trim();
                        string HS04 = line.Substring(1144, 12).Trim();
                        string A4114 = line.Substring(1156, 12).Trim();
                        string HS05 = line.Substring(1168, 12).Trim();
                        string A4115 = line.Substring(1180, 12).Trim();
                        string HS06 = line.Substring(1192, 12).Trim();
                        string A4116 = line.Substring(1204, 12).Trim();
                        string HS07 = line.Substring(1216, 12).Trim();
                        string A4117 = line.Substring(1228, 12).Trim();
                        string A4172 = line.Substring(1240, 12).Trim();
                        string A4173 = line.Substring(1252, 12).Trim();
                        string HC02 = line.Substring(1264, 12).Trim();
                        string A4464 = line.Substring(1276, 12).Trim();
                        string A5100 = line.Substring(1288, 12).Trim();
                        string A5700 = line.Substring(1300, 12).Trim();
                        string A5702 = line.Substring(1312, 12).Trim();
                        string A5220 = line.Substring(1324, 12).Trim();
                        string A5400 = line.Substring(1336, 12).Trim();
                        string A5800 = line.Substring(1348, 12).Trim();
                        string A5900 = line.Substring(1360, 12).Trim();
                        string A7111 = line.Substring(1372, 12).Trim();
                        string A7114 = line.Substring(1384, 12).Trim();
                        string A7112 = line.Substring(1396, 12).Trim();
                        string A7116 = line.Substring(1408, 12).Trim();
                        string A7117 = line.Substring(1420, 12).Trim();
                        string A7113 = line.Substring(1432, 12).Trim();
                        string A5920 = line.Substring(1444, 12).Trim();
                        string A59209 = line.Substring(1456, 12).Trim();
                        string A7115 = line.Substring(1468, 12).Trim();
                        string A7600 = line.Substring(1480, 12).Trim();
                        string A7670 = line.Substring(1492, 12).Trim();
                        string A3910 = line.Substring(1504, 12).Trim();
                        string A4170 = line.Substring(1516, 12).Trim();
                        string A4811 = line.Substring(1528, 12).Trim();
                        string A4813 = line.Substring(1540, 12).Trim();
                        string A4812 = line.Substring(1552, 12).Trim();
                        string A4830 = line.Substring(1565, 12).Trim();
                        string BRUT = line.Substring(1576, 12).Trim();
                        string A9110 = line.Substring(1588, 12).Trim();
                        string A9112 = line.Substring(1600, 12).Trim();
                        string A6110 = line.Substring(1612, 12).Trim();
                        string A6120 = line.Substring(1624, 12).Trim();
                        string A9121 = line.Substring(1636, 12).Trim();
                        string A6310 = line.Substring(1648, 12).Trim();
                        string A6312 = line.Substring(1660, 12).Trim();
                        string A6311 = line.Substring(1672, 12).Trim();
                        string A6462 = line.Substring(1684, 12).Trim();
                        string A6511 = line.Substring(1696, 12).Trim();
                        string A6430 = line.Substring(1709, 12).Trim();
                        string A6465 = line.Substring(1720, 12).Trim();
                        string A6467 = line.Substring(1732, 12).Trim();
                        string A6840 = line.Substring(1744, 12).Trim();
                        string A6810 = line.Substring(1756, 12).Trim();
                        string A6830 = line.Substring(1768, 12).Trim();
                        string A6860 = line.Substring(1780, 12).Trim();
                        string A6463 = line.Substring(1792, 12).Trim();
                        string A6610 = line.Substring(1804, 12).Trim();
                        string A6870 = line.Substring(1816, 12).Trim();
                        string A8140 = line.Substring(1828, 12).Trim();
                        string RUB_DEDSA = line.Substring(1840, 12).Trim();
                        string NETPAIE = line.Substring(1852, 12).Trim();
                        string A61109 = line.Substring(1865, 12).Trim();
                        string A61209 = line.Substring(1876, 12).Trim();
                        string A6150 = line.Substring(1888, 12).Trim();
                        string A6210 = line.Substring(1900, 12).Trim();
                        string TFP = line.Substring(1912, 12).Trim();
                        string RUB_COTEPY = line.Substring(1924, 12).Trim();
                        string RD_NET = line.Substring(1936, 12).Trim();
                        string MoisPaie = line.Substring(1948, 9).Trim();
                        string MoisEnt = line.Substring(1957, 2).Trim();
                        string RIB = line.Substring(1959, 20).Trim();
                        string CUMNB2 = line.Substring(1979, 12).Trim();
                        prog.Value = prog.Value + 1;
                        int Annee = DateTime.Today.Year;
                        int Mois = int.Parse(MoisEnt);
                        service.insertGrossToNetA016(matricule, nom.Replace("�", "e"), prenom.Replace("�", "e"),
                            DateEmbauche, DateDept, CodeDept, IntituleDept.Replace("�", "e"),
                        NatureContrat, Emploi.Replace("�", "e"), SALBASE, HORAIRE,A1000,A2100,A2200,A3413,A3414,A3415,
                        A3418,A3316,A3317,A3510,A3520,A3530,A3417,A5910,A3318,A3610,A3620,A3630,A3640,A3651,A3652,A3650,
                        A3660,A3661,A3670,A3680,A3681,A3682,A3683,A3684,A3685,BRUTFX,A3686,A3690,CL31,A4460,CL32,A4461,
                        CL33,A4462,CL34,A4463,CL35,A3110,CL37,A4292,A4293,CL39,A4294,HC01,A4295,CL40,A4296,A4491,A4162,
                        A4391,A4381,A4840,A4382,A4850,A4383,A4860,A4718,A4717,A4713,A4716,HS01,A4111,HS02,A4112,HS03,
                        A4113,HS04,A4114,HS05,A4115,HS06,A4116,HS07,A4117,A4172,A4173,HC02,A4464,A5100,A5700,A5702,A5220,
                        A5400,A5800,A5900,A7111,A7114,A7112,A7116,A7117,A7113,A5920,A59209,A7115,A7600,A7670,A3910,A4170,
                        A4811,A4813,A4812,A4830,BRUT,A9110,A9112,A6110,A6120,A9121,A6310,A6312,A6311,A6462,A6511,A6430,A6465,
                        A6467, A6840,A6810,A6830,A6860,A6463,A6610,A6870,A8140,RUB_DEDSA,NETPAIE,A61109,A61209,A6150,A6210,
                        TFP, RUB_COTEPY, RD_NET, MoisPaie, NumSoc, Mois, Annee, RIB, CUMNB2);
                    }
                }

            }
            MessageBox.Show("Import realized with success");
            //       prog.Value = 0;
            //       prog.Visible = false;
        }

        public void EcrireDataReport(int NumSoc, string file, ProgressBar prog)
        {
             StreamReader sr = new StreamReader(file);
            string line = null;
            prog.Visible = true;
            prog.Minimum = 0;
            bool rt = false;
            int ligne = 0;
            while ((line = sr.ReadLine()) != null)
            {
                ligne = ligne + 1;
                prog.Maximum = ligne;
                if (!(line.Equals("")))
                {
                    string matricule = line.Substring(0, 10).Trim();
                    string nom = line.Substring(10, 80).Trim().Replace("'", "''");
                    string prenom = line.Substring(90, 20).Trim().Replace("'", "''");
                    string DateEmbauche = line.Substring(110, 8).Trim();
                    string DateDept = line.Substring(118, 8).Trim();
                    string ModePe = line.Substring(126, 10).Trim();
                    string InfoBank = line.Substring(136, 24).Trim().Replace("'", "''");
                    string NetPaie = line.Substring(160, 12).Trim();
                    string Mois = line.Substring(172, 2).Trim();
                    prog.Value = prog.Value + 1;
                    int Annee = DateTime.Today.Year;
                    if ((matricule != "") && (nom != "") && (prenom != ""))
                    {
                       rt = service.insertPaymentReport(matricule, nom, prenom, DateEmbauche, DateDept, ModePe, InfoBank, NetPaie, NumSoc, int.Parse(Mois), Annee);
                    }
             /*       if (rt)
                    {
                        MessageBox.Show("Add realized with success");
                    }
                    else
                    {
                        MessageBox.Show("Error to add this DATA", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }*/
                }
            }
            MessageBox.Show("Import realized with success");

        }

        public string EmplacementFich()
        {
            int taille = 0;
            string EmpFinal = string.Empty;
            string file = string.Empty;
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Title = "Choisir le fichier à importer";
            // Extension par défaut			
            openFileDialog1.DefaultExt = "txt";
            // Filtre sur les fichiers
            openFileDialog1.Filter = "Tous les fichiers (*.txt)|*.txt";
            openFileDialog1.Multiselect = false;
            openFileDialog1.FilterIndex = 1;
            // Ouverture boite de dialogue OpenFile
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                file = openFileDialog1.FileName;
                taille = file.Length - 3;
                if ((file.Substring(taille, 3) == "txt") || (file.Substring(taille, 3) == "TXT"))
                {
                    EmpFinal = file;
                }
                else
                {
                    EmpFinal = null;
                }
            }
            return EmpFinal;
        }

        public string EmplacementExcelFile()
        {
            int taille = 0;
            string EmpFinal = string.Empty;
            string file = string.Empty;
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Title = "Choisir le fichier à importer";
            // Extension par défaut			
            openFileDialog1.DefaultExt = "xls";
            // Filtre sur les fichiers
            openFileDialog1.Filter = "Tous les fichiers (*.xls)|*.xls";
            openFileDialog1.Multiselect = false;
            openFileDialog1.FilterIndex = 1;
            // Ouverture boite de dialogue OpenFile
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                file = openFileDialog1.FileName;
                taille = file.Length - 3;
                if ((file.Substring(taille, 3) == "xls") || (file.Substring(taille, 3) == "xlsx"))
                {
                    EmpFinal = file;
                }
                else
                {
                    EmpFinal = null;
                }
            }
            return EmpFinal;
        }

        public bool EcrireDataFixe(int NumSoc, string file,ProgressBar prog)
        {
            prog.Visible = true;
            prog.Minimum = 0;
            prog.Maximum = 2500;
            bool verif = false;
            StreamReader sr = new StreamReader(file);
            string line = null;
            int ligne = 0;
            while ((line = sr.ReadLine()) != null)
            {
                ligne = ligne + 1;

                if (!(line.Equals("")))
                {
                    string IdOracle = line.Substring(0, 10).Trim();
                    string IdLocale = line.Substring(10, 10).Trim();
                    string Civilite = line.Substring(20, 12).Trim();
                    string CiviliteAbrege = line.Substring(32, 4).Trim();
                    string CiviliteValeur = line.Substring(36, 1).Trim();
                    string Sexe = line.Substring(37, 8).Trim();
                    string SexeVal = line.Substring(45, 1).Trim();
                    string Nom = line.Substring(46, 80).Trim().Replace("'", "''");
                    string NomJFille = line.Substring(126, 80).Trim().Replace("'", "''");
                    string Prenom = line.Substring(206, 20).Trim().Replace("'", "''");
                    string Adresse = line.Substring(226, 32).Trim().Replace("'", "''");
                    string ComplAdr = line.Substring(258, 32).Trim().Replace("'", "''");
                    string Commune = line.Substring(290, 26).Trim();
                    string CodePostal = line.Substring(316, 5).Trim();
                    string Tel = line.Substring(321, 15).Trim();
                    string CodePays = line.Substring(336, 3).Trim();
                    string InitulePays = line.Substring(339, 30).Trim().Replace("'", "''");
                    string Nationalite = line.Substring(369, 3).Trim();
                    string IntituleNationalite = line.Substring(372, 40).Trim().Replace("'", "''");
                    string NumCarteSejour = line.Substring(412, 10).Trim();
                    string DateExpCartSejour = line.Substring(422, 8).Trim();
                    string CodePaysNaiss = line.Substring(430, 3).Trim();
                    string DateNaiss = line.Substring(433, 8).Trim();
                    string CommuneNaiss = line.Substring(441, 26).Trim().Replace("'", "''");
                    string SitFamiliale = line.Substring(467, 20).Trim();
                    string SitFamEnValeur = line.Substring(487, 1).Trim();
                    string NbreEnf = line.Substring(488, 2).Trim();
                    string EnfRenseigne = line.Substring(490, 2).Trim();
                    string NumCNSS = line.Substring(492, 13).Trim();
                    string CleCNSS = line.Substring(505, 2).Trim();
                    string DateEmbSoc = line.Substring(507, 8).Trim();
                    string DateDepSoc = line.Substring(515, 8).Trim();
                    string EtabSal = line.Substring(523, 5).Trim().Replace("'", "''"); ;
                    string TypeEntEtab = line.Substring(528, 2).Trim();
                    string DateSorEtab = line.Substring(530, 8).Trim();
                    string CodeConvColleSal = line.Substring(538, 3).Trim();
                    string DeptSal = line.Substring(541, 10).Trim().Replace("'", "''");
                    string IntituleDept = line.Substring(551, 30).Trim().Replace("'", "''");
                    string SceSal = line.Substring(581, 10).Trim().Replace("'", "''");
                    string UniteSal = line.Substring(591, 10).Trim();
                    string IntituleUnite = line.Substring(601, 30).Trim().Replace("'", "''");
                    string CategSal = line.Substring(631, 10).Trim();
                    string CodeFct = line.Substring(641, 12).Trim();
                    string EmploiOcc = line.Substring(653, 30).Trim().Replace("'", "''");
                    string NatureContrat = line.Substring(683, 10).Trim();
                    string DateDebContrat = line.Substring(693, 8).Trim();
                    string BullModSal = line.Substring(701, 4).Trim();
                    string TypeSal = line.Substring(705, 18).Trim();
                    string SalBaseSal = line.Substring(723, 12).Trim();
                    string HorBaseSal = line.Substring(735, 12).Trim();
                    string CongRestAnnePrec = line.Substring(747, 13).Trim();
                    string ModePe = line.Substring(760, 10).Trim();
                    string CodeBque1 = line.Substring(770, 5).Trim();
                    string CodeGuichet1 = line.Substring(775, 5).Trim();
                    string NumCpt1 = line.Substring(780, 11).Trim();
                    string CleRib1 = line.Substring(791, 2).Trim();
                    string NomGuichet1 = line.Substring(793, 30).Trim();
                    string NomBque1 = line.Substring(823, 30).Trim();
                    string LibCpt1 = line.Substring(853, 24).Trim().Replace("'", "''");
                    string MtAssVie = line.Substring(877, 16).Trim();
                    string EnfACharge = line.Substring(893, 12).Trim();
                    string ChefFamille = line.Substring(905, 12).Trim();
                    string AdhesionASGroup = line.Substring(917, 12).Trim();
                    string ChampAlfa1 = line.Substring(929, 20).Trim();
                    string ChampAlfa2 = line.Substring(949, 20).Trim();
                    string ChampAlfa3 = line.Substring(969, 20).Trim();
                    string ChampAlfa4 = line.Substring(989, 10).Trim();
                    string ChampAlfa5 = line.Substring(999, 10).Trim();
                    string ChampAlfa6 = line.Substring(1009, 10).Trim();
                    string RtCalcule = line.Substring(1019, 12).Trim();
                    string CodeSce = line.Substring(1031, 10).Trim();
                    string IntituleSce = line.Substring(1041, 30).Trim().Replace("'", "''"); ;
                    string Date1 = line.Substring(1071, 8).Trim();
                    prog.Value = prog.Value + 1;

                    service.insertDataFixe(IdOracle, IdLocale, Civilite, CiviliteAbrege, CiviliteValeur, Sexe, SexeVal, Nom.Replace("�", "e"), NomJFille.Replace("�", "e"), Prenom.Replace("�", "e"),
                        Adresse.Replace("�", "e"), ComplAdr.Replace("�", "e"), Commune, CodePostal, Tel, CodePays, InitulePays.Replace("�", "e"), Nationalite, IntituleNationalite.Replace("�", "e"), NumCarteSejour,
                        DateExpCartSejour, CodePaysNaiss, DateNaiss, CommuneNaiss.Replace("�", "e"), SitFamiliale, SitFamEnValeur, NbreEnf, EnfRenseigne, NumCNSS,
                        CleCNSS, DateEmbSoc, DateDepSoc, EtabSal, TypeEntEtab, DateSorEtab, CodeConvColleSal, DeptSal, IntituleDept.Replace("�", "e"), SceSal,
                        UniteSal, IntituleUnite.Replace("�", "e"), CategSal, CodeFct, EmploiOcc.Replace("�", "e"),NatureContrat, DateDebContrat, BullModSal, TypeSal, SalBaseSal, HorBaseSal, CongRestAnnePrec,
                        ModePe, CodeBque1, CodeGuichet1, NumCpt1, CleRib1, NomGuichet1, NomBque1, LibCpt1, MtAssVie, EnfACharge, ChefFamille, AdhesionASGroup,
                        ChampAlfa1, ChampAlfa2, ChampAlfa3, ChampAlfa4, ChampAlfa5, ChampAlfa6, RtCalcule, CodeSce, IntituleSce.Replace("�", "e"), Date1, NumSoc);
                    verif = true;
                }
                else
                {
                    prog.Maximum = ligne;
                    prog.Value = 0;
                    prog.Visible = false;
                    verif = false;

                }
            }
            MessageBox.Show("Import realized with success");
            return verif;
        }

        public void ExtraireDataSRF(string Date1,string Date2,string CodeSoc,string MoisPe,string NomSoc,string DateDeb,string DateFin,
            bool PeNormal,bool Bonus,string DatePe,int idSoc,ProgressBar prog,string Annee)
        {
            string nomFichier = string.Empty;
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.DefaultExt = "xls";
            saveFileDialog1.FilterIndex = 2;
            saveFileDialog1.Filter = "fichiers excel (*.xls)|*.xls";
            saveFileDialog1.RestoreDirectory = true;
            saveFileDialog1.FileName = "P_SRF_"+CodeSoc+"_"+Date1+"_"+Date2+"_"+MoisPe+".xls.xls";
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                try
                {

                    nomFichier = saveFileDialog1.FileName;
                    EcrireData(nomFichier, CodeSoc, NomSoc, DateDeb, DateFin,PeNormal,Bonus,DatePe,idSoc,prog,Annee,MoisPe);
                }
                catch (Exception ex)
                {
                    ex.ToString();
                }
            }
        }
        public Worksheet EcritureC(int refSoc, ProgressBar prog, string MoisPe, string Annee, Worksheet worksheet, string NomSoc,string CodeSoc)
       {
           string Nom = string.Empty; string ModePaie = string.Empty; string Prenom = string.Empty;
           string MatSal = string.Empty; string NomJeuneFille = string.Empty; string CIN = string.Empty;
           string CivilAbr = string.Empty; string DateNaiss = string.Empty; string CodePaysNaiss = string.Empty;
           string Gender = string.Empty; string sexe = string.Empty; string CivilVal = string.Empty;
           string Adr = string.Empty; string ComplAdr = string.Empty; string Commune = string.Empty;
           string CodeP = string.Empty; string CodeBank, CodeGuichet, NumCpt, CleRib = string.Empty;
           string RIB = string.Empty; string NomBque = string.Empty; string Avance = string.Empty;
           double Difference = 0.0; string Emploi = string.Empty; string NatContrat = string.Empty; string NbreEnfant = string.Empty;
           string CodeSce = string.Empty; string DateEmb = string.Empty; double GSL_BSS = 0.0; string SitFam = string.Empty;
           string IndemFct = string.Empty; double GSL_SBC = 0.0; string PrAstreinte2 = string.Empty; string SitFamVal = string.Empty;
           double GSL_OTC = 0.0; double GSL_LEAV = 0.0; string MatLoc = string.Empty; string LaSoc = string.Empty;
           string LibCpte = string.Empty; string DateSortie = string.Empty; string DeptSal = string.Empty; string IntituleDept = string.Empty;
           double GSL_NTHC = 0.0; double GSL_BNS = 0.0; double GSL_COM = 0.0; double GSL_OVS = 0.0;
           double GSL_OBIK = 0.0; double GSL_TALW = 0.0; double GSL_CALW = 0.0; double GSL_SMS = 0.0;
           double GSL_HSD = 0.0; double GSL_HSC = 0.0; double GSL_LSD = 0.0; double GSL_LSC = 0.0;
           double GSL_CCYN = 0.0; double GSL_CCBIK = 0.0; double GSL_HALW = 0.0; double GSL_MALW = 0.0;
           double GSL_OALW = 0.0; double GSL_FEX = 0.0; double GSL_TPP = 0.0; double EED_WTX = 0.0;
           double EED_CSS = 0.0; double EED_CRT = 0.0; double EED_CUN = 0.0; double EED_CMC = 0.0;
           double NTP_SPPD = 0.0; double NTP_EXRF = 0.0; double NTP_EXPA = 0.0; double NTP_BIK = 0.0;
           double EED_VOTH = 0.0; double NTP_ONAD = 0.0; double NTP_SAA = 0.0; double NTP_TPP = 0.0;
           double EED_VMC = 0.0; double ERC_WTX = 0.0; double ERC_CSS = 0.0; double ERC_CRT = 0.0;
           double EED_VRT = 0.0; double ERC_CUN = 0.0; double ERC_CMC = 0.0; double ERC_MCC = 0.0;
           double EED_COTH = 0.0; double ERC_AWI = 0.0; double ERC_MCI = 0.0; double ERC_MTX = 0.0;
           double ERC_VOTH = 0.0; double ACR_HLS = 0.0; double ACR_HLC = 0.0; double ACR_BNS = 0.0;
           double ACR_BNC = 0.0; double ACR_SMS = 0.0; double ACR_SMC = 0.0; double EXT_CCBIK = 0.0;
           double EXT_OBIK = 0.0;
           string SalNet = string.Empty;
           string HSupp = string.Empty;             
               int i = 1;
               DataTable df = service.RecupererData(refSoc);
               foreach (System.Data.DataRow re in df.Rows)
               {
                       prog.Value = prog.Value + 1;
                       ModePaie = re["ModePaie"].ToString();
                       LaSoc = NomSoc;
                       MatSal = re["IDOracle"].ToString();
                       MatLoc = re["IDLocale"].ToString();
                       Nom = re["Nom"].ToString();
                       NomJeuneFille = re["NomJeuneFille"].ToString();
                       Prenom = re["Prenom"].ToString();
                       CIN = re["CIN"].ToString();
                       SitFam = re["SitFam"].ToString();
                       SitFamVal = re["SitFamVal"].ToString();
                       NbreEnfant = re["EnfRenseignes"].ToString();
                       CivilAbr = re["CiviliteAbrege"].ToString();
                       DateNaiss = DateMank(re["DateNaiss"].ToString());
                       CodePaysNaiss = re["CodePays"].ToString();
                       Gender = re["SexeValeur"].ToString();
                       sexe = re["Sexe"].ToString();
                       CivilVal = re["CiviliteValeur"].ToString();
                       LibCpte = re["LibCpte"].ToString();
                       Adr = re["Adresse"].ToString();
                       ComplAdr = re["ComplAdresse"].ToString();
                       Commune = re["Commune"].ToString();
                       CodeP = VerifSiVide(re["CodePostal"].ToString());
                       if (re["CodeBque"].ToString() == "")
                       {
                           CodeBank = VerifSiVide(re["CodeBque"].ToString());
                       }
                       else
                       {
                           CodeBank = VerifSiVide(re["CodeBque"].ToString().Substring(3, 2));
                       }
                      
                       CodeGuichet = VerifSiVide(re["CodeGuichet"].ToString());
                       NumCpt = VerifSiVide(re["NumCpt"].ToString());
                       CleRib = VerifSiVide(re["CleRib"].ToString());
                       RIB = CodeBank + CodeGuichet + NumCpt + CleRib;
                       NomBque = re["NomBque"].ToString();
                       Emploi = re["EmploiOccupee"].ToString();
                       NatContrat = re["NatureContrat"].ToString();
                       DateEmb = DateMank(re["DateEmbauche"].ToString());
                       DateSortie = DateMank(re["DateSortieEtab"].ToString());
                       DeptSal = re["DeptSal"].ToString();
                       IntituleDept = re["IntituleDept"].ToString();
                       CodeSce = VerifSiVide(re["CodeSce"].ToString());
                       GSL_BSS = service.DataCodeMemo("123", MatSal);
                       GSL_SMS = service.DataCodeMemo("124", MatSal);
                       GSL_SBC = service.DataCodeMemo("125", MatSal);
                       GSL_OTC = service.DataCodeMemo("126", MatSal);
                       GSL_BNS = service.DataCodeMemo("127", MatSal);
                       GSL_COM = service.DataCodeMemo("128", MatSal);
                       GSL_OVS = service.DataCodeMemo("129", MatSal);
                       GSL_LEAV = service.DataCodeMemo("130", MatSal);
                       GSL_HSD = service.DataCodeMemo("131", MatSal);
                       GSL_HSC = service.DataCodeMemo("132", MatSal);
                       GSL_LSD = service.DataCodeMemo("133", MatSal);
                       GSL_LSC = service.DataCodeMemo("134", MatSal);
                       GSL_NTHC = service.DataCodeMemo("135", MatSal);
                       GSL_CCYN = service.DataCodeMemo("136", MatSal);
                       GSL_CCBIK = service.DataCodeMemo("137", MatSal);
                       GSL_OBIK = service.DataCodeMemo("138", MatSal);
                       GSL_HALW = service.DataCodeMemo("139", MatSal);
                       GSL_MALW = service.DataCodeMemo("140", MatSal);
                       GSL_TALW = service.DataCodeMemo("141", MatSal);
                       GSL_CALW = service.DataCodeMemo("142", MatSal);
                       GSL_OALW = service.DataCodeMemo("143", MatSal);
                       GSL_FEX = service.DataCodeMemo("144", MatSal);
                       GSL_TPP = service.DataCodeMemo("145", MatSal);
                       EED_WTX = service.DataCodeMemo("147", MatSal);
                       EED_CSS = service.DataCodeMemo("148", MatSal);
                       EED_CRT = service.DataCodeMemo("149", MatSal);
                       EED_CUN = service.DataCodeMemo("150", MatSal);
                       EED_CMC = service.DataCodeMemo("151", MatSal);
                       EED_COTH = service.DataCodeMemo("152", MatSal);
                       EED_VRT = service.DataCodeMemo("153", MatSal);
                       EED_VMC = service.DataCodeMemo("154", MatSal);
                       EED_VOTH = service.DataCodeMemo("155", MatSal);
                       NTP_SPPD = service.DataCodeMemo("158", MatSal);
                       NTP_EXRF = service.DataCodeMemo("159", MatSal);
                       NTP_EXPA = service.DataCodeMemo("160", MatSal);
                       NTP_BIK = service.DataCodeMemo("161", MatSal);
                       NTP_ONAD = service.DataCodeMemo("162", MatSal);
                       NTP_SAA = service.DataCodeMemo("163", MatSal);
                       NTP_TPP = service.DataCodeMemo("164", MatSal);
                       ERC_WTX = service.DataCodeMemo("167", MatSal);
                       ERC_CSS = service.DataCodeMemo("400", MatSal);
                       ERC_CRT = service.DataCodeMemo("169", MatSal);
                       ERC_CUN = service.DataCodeMemo("170", MatSal);
                       ERC_CMC = service.DataCodeMemo("171", MatSal);
                       ERC_MCC = service.DataCodeMemo("172", MatSal);
                       ERC_AWI = service.DataCodeMemo("173", MatSal);
                       ERC_MCI = service.DataCodeMemo("174", MatSal);
                       ERC_MTX = service.DataCodeMemo("175", MatSal);
                       ERC_VOTH = service.DataCodeMemo("176", MatSal);
                       ACR_HLS = service.DataCodeMemo("178", MatSal);
                       ACR_HLC = service.DataCodeMemo("179", MatSal);
                       ACR_BNS = service.DataCodeMemo("180", MatSal);
                       ACR_BNC = service.DataCodeMemo("181", MatSal);
                       ACR_SMS = service.DataCodeMemo("182", MatSal);
                       ACR_SMC = service.DataCodeMemo("183", MatSal);
                       EXT_CCBIK = service.DataCodeMemo("185", MatSal);
                       EXT_OBIK = service.DataCodeMemo("186", MatSal);


                       if (refSoc == 2)
                       {
                           SalNet = service.SalNet(MatSal, A001, int.Parse(MoisPe), int.Parse(Annee));
                           Avance = service.Avance(MatSal, A001, int.Parse(MoisPe), int.Parse(Annee));
                           double Net = VerifierSiNull(SalNet);
                           double Av = VerifierSiNull(Avance);
                           Difference = Net - Av;

                       }
                       if (refSoc == 5)
                       {
                           SalNet = service.SalNet(MatSal, A002, int.Parse(MoisPe), int.Parse(Annee));
                           Avance = service.Avance(MatSal, A002, int.Parse(MoisPe), int.Parse(Annee));
                           double Net = VerifierSiNull(SalNet);
                           double Av = VerifierSiNull(Avance);
                           Difference = Net - Av;
                       }
                       if (refSoc == 8)
                       {
                           SalNet = service.SalNet(MatSal, A003, int.Parse(MoisPe), int.Parse(Annee));
                           Avance = service.Avance(MatSal, A003, int.Parse(MoisPe), int.Parse(Annee));
                           double Net = VerifierSiNull(SalNet);
                           double Av = VerifierSiNull(Avance);
                           Difference = Net - Av;
                       }
                       if (refSoc == 7)
                       {
                           SalNet = service.SalNet(MatSal, A004, int.Parse(MoisPe), int.Parse(Annee));
                           Avance = service.Avance(MatSal, A004, int.Parse(MoisPe), int.Parse(Annee));
                           double Net = VerifierSiNull(SalNet);
                           double Av = VerifierSiNull(Avance);
                           Difference = Net - Av;
                       }
                       if (refSoc == 9)
                       {
                           SalNet = service.SalNet(MatSal, A010, int.Parse(MoisPe), int.Parse(Annee));
                           Avance = service.Avance(MatSal, A010, int.Parse(MoisPe), int.Parse(Annee));
                           double Net = VerifierSiNull(SalNet);
                           double Av = VerifierSiNull(Avance);
                           Difference = Net - Av;
                       }
                       if (refSoc == 10)
                       {
                           SalNet = service.SalNet(MatSal, A011, int.Parse(MoisPe), int.Parse(Annee));
                           Avance = service.Avance(MatSal, A011, int.Parse(MoisPe), int.Parse(Annee));
                           double Net = VerifierSiNull(SalNet);
                           double Av = VerifierSiNull(Avance);
                           Difference = Net - Av;
                       }
                       if (refSoc == 11)
                       {
                           SalNet = service.SalNet(MatSal, A013, int.Parse(MoisPe), int.Parse(Annee));
                           Avance = service.Avance(MatSal, A013, int.Parse(MoisPe), int.Parse(Annee));
                           double Net = VerifierSiNull(SalNet);
                           double Av = VerifierSiNull(Avance);
                           Difference = Net - Av;
                       }
                       if (refSoc == 3)
                       {
                           SalNet = service.SalNet(MatSal, A014, int.Parse(MoisPe), int.Parse(Annee));
                           Avance = service.Avance(MatSal, A014, int.Parse(MoisPe), int.Parse(Annee));
                           double Net = VerifierSiNull(SalNet);
                           double Av = VerifierSiNull(Avance);
                           Difference = Net - Av;
                       }
                       if (refSoc == 12)
                       {
                           SalNet = service.SalNet(MatSal, A015, int.Parse(MoisPe), int.Parse(Annee));
                           Avance = service.Avance(MatSal, A015, int.Parse(MoisPe), int.Parse(Annee));
                           double Net = VerifierSiNull(SalNet);
                           double Av = VerifierSiNull(Avance);
                           Difference = Net - Av;
                       }

                       if (refSoc == 4)
                       {
                           SalNet = service.SalNet(MatSal, A020, int.Parse(MoisPe), int.Parse(Annee));
                           Avance = service.Avance(MatSal, A020, int.Parse(MoisPe), int.Parse(Annee));
                           double Net = VerifierSiNull(SalNet);
                           double Av = VerifierSiNull(Avance);
                           Difference = Net - Av;
                       }
                       worksheet = entete.EnteteSRFLigne4(worksheet,refSoc, ModePaie,MatLoc,MatSal, Nom, NomJeuneFille, Prenom, CIN, CivilAbr, DateNaiss,
                      CodePaysNaiss, Gender.ToString(), sexe, CivilVal, Adr, ComplAdr, Commune, CodeP,SitFam,SitFamVal,NbreEnfant,LibCpte, RIB, NomBque, Difference.ToString("#.##"),
                      Emploi, NatContrat, DateEmb,DateSortie,DeptSal,IntituleDept, NomSoc, CodeSce, GSL_BSS, GSL_SMS, GSL_SBC, GSL_OTC, GSL_BNS, GSL_COM, GSL_OVS, GSL_LEAV,
                      GSL_HSD, GSL_HSC, GSL_LSD, GSL_LSC, GSL_NTHC, GSL_CCYN, GSL_CCBIK, GSL_OBIK,
                      GSL_HALW, GSL_MALW, GSL_TALW, GSL_CALW, GSL_OALW, GSL_FEX, GSL_TPP, EED_WTX,
                      EED_CSS, EED_CRT, EED_CUN, EED_CMC, EED_COTH, EED_VRT, EED_VMC, EED_VOTH, NTP_SPPD,
                      NTP_EXRF, NTP_EXPA, NTP_BIK, NTP_ONAD, NTP_SAA, NTP_TPP, ERC_WTX, ERC_CSS,
                      ERC_CRT, ERC_CUN, ERC_CMC, ERC_MCC, ERC_AWI, ERC_MCI, ERC_MTX, ERC_VOTH, ACR_HLS,
                      ACR_HLC, ACR_BNS, ACR_BNC, ACR_SMS, ACR_SMC, EXT_CCBIK,i,CodeSoc);
                       i = i + 1;
                   }
              
           return worksheet;
       }

        public string DateMank(string DateF)
        {
            string DateR = string.Empty;
            if (DateF == "")
            {
                DateR = "Vide";
            }
            else
            {
                DateR = DateF;
            }
            return DateR;
        }

        public string VerifSiVide(string Valeur)
        {
            string ValF = string.Empty;
            if (Valeur == "")
            {
                ValF = "0";
            }
            else
            {
                ValF = Valeur;
            }
            return ValF;
        }

        public Worksheet EcritFooter(Worksheet worksheet)
        {
           double GSL_BSS = service.SUMCodeMemo("123");
           double GSL_SMS = service.SUMCodeMemo("124");
           double GSL_SBC = service.SUMCodeMemo("125");
           double GSL_OTC = service.SUMCodeMemo("126");
           double GSL_BNS = service.SUMCodeMemo("127");
           double GSL_COM = service.SUMCodeMemo("128");
           double GSL_OVS = service.SUMCodeMemo("129");
           double GSL_LEAV = service.SUMCodeMemo("130");
           double GSL_HSD = service.SUMCodeMemo("131");
           double GSL_HSC = service.SUMCodeMemo("132");
           double GSL_LSD = service.SUMCodeMemo("133");
           double GSL_LSC = service.SUMCodeMemo("134");
           double GSL_NTHC = service.SUMCodeMemo("135");
           double GSL_CCYN = service.SUMCodeMemo("136");
           double GSL_CCBIK = service.SUMCodeMemo("137");
           double GSL_OBIK = service.SUMCodeMemo("138");
           double GSL_HALW = service.SUMCodeMemo("139");
           double GSL_MALW = service.SUMCodeMemo("140");
           double GSL_TALW = service.SUMCodeMemo("141");
           double GSL_CALW = service.SUMCodeMemo("142");
           double GSL_OALW = service.SUMCodeMemo("143");
           double GSL_FEX = service.SUMCodeMemo("144");
           double GSL_TPP = service.SUMCodeMemo("145");
           double EED_WTX = service.SUMCodeMemo("147");
           double EED_CSS = service.SUMCodeMemo("148");
           double EED_CRT = service.SUMCodeMemo("149");
           double EED_CUN = service.SUMCodeMemo("150");
           double EED_CMC = service.SUMCodeMemo("151");
           double EED_COTH = service.SUMCodeMemo("152");
           double EED_VRT = service.SUMCodeMemo("153");
           double EED_VMC = service.SUMCodeMemo("154");
           double EED_VOTH = service.SUMCodeMemo("155");
           double NTP_SPPD = service.SUMCodeMemo("158");
           double NTP_EXRF = service.SUMCodeMemo("159");
           double NTP_EXPA = service.SUMCodeMemo("160");
           double NTP_BIK = service.SUMCodeMemo("161");
           double NTP_ONAD = service.SUMCodeMemo("162");
           double NTP_SAA = service.SUMCodeMemo("163");
           double NTP_TPP = service.SUMCodeMemo("164");
           double ERC_WTX = service.SUMCodeMemo("167");
           double ERC_CSS = service.SUMCodeMemo("400");
           double ERC_CRT = service.SUMCodeMemo("169");
           double ERC_CUN = service.SUMCodeMemo("170");
           double ERC_CMC = service.SUMCodeMemo("171");
           double ERC_MCC = service.SUMCodeMemo("172");
           double ERC_AWI = service.SUMCodeMemo("173");
           double ERC_MCI = service.SUMCodeMemo("174");
           double ERC_MTX = service.SUMCodeMemo("175");
           double ERC_VOTH = service.SUMCodeMemo("176");
           double ACR_HLS = service.SUMCodeMemo("178");
           double ACR_HLC = service.SUMCodeMemo("179");
           double ACR_BNS = service.SUMCodeMemo("180");
           double ACR_BNC = service.SUMCodeMemo("181");
           double ACR_SMS = service.SUMCodeMemo("182");
           double ACR_SMC = service.SUMCodeMemo("183");
           double EXT_CCBIK = service.SUMCodeMemo("185");
           double EXT_OBIK = service.SUMCodeMemo("186");
           int i = service.NbSal();
           worksheet = entete.EnteteSRFLigneFooter(worksheet, GSL_BSS, GSL_SMS, GSL_SBC, GSL_OTC, GSL_BNS, GSL_COM, GSL_OVS, GSL_LEAV,
                    GSL_HSD, GSL_HSC, GSL_LSD, GSL_LSC, GSL_NTHC, GSL_CCYN, GSL_CCBIK, GSL_OBIK,
                    GSL_HALW, GSL_MALW, GSL_TALW, GSL_CALW, GSL_OALW, GSL_FEX, GSL_TPP, EED_WTX,
                    EED_CSS, EED_CRT, EED_CUN, EED_CMC, EED_COTH, EED_VRT, EED_VMC, EED_VOTH, NTP_SPPD,
                    NTP_EXRF, NTP_EXPA, NTP_BIK, NTP_ONAD, NTP_SAA, NTP_TPP, ERC_WTX, ERC_CSS,
                    ERC_CRT, ERC_CUN, ERC_CMC, ERC_MCC, ERC_AWI, ERC_MCI, ERC_MTX, ERC_VOTH, ACR_HLS,
                    ACR_HLC, ACR_BNS, ACR_BNC, ACR_SMS, ACR_SMC, EXT_CCBIK, i);
           return worksheet;
        }

        public void EcrireData(string nomFichier, string CodeSoc, string NomSoc, string Date1, string Date2, bool PeNormal, bool Bonus,
            string DatePe, int refSoc, ProgressBar prog, string Annee, string MoisPe)
        {
            
            prog.Visible = true;
            prog.Minimum = 0;
            //create new xls file
            Workbook workbook = new Workbook();
            Worksheet worksheet = new Worksheet("SRF");
            worksheet.Cells.ColumnWidth[0, 1] = 3000;
            worksheet = entete.EnteteSRFLigne1(worksheet);
            worksheet = entete.EnteteSRFLigne2(worksheet, NomSoc, Date1, Date2, CodeSoc, PeNormal, Bonus, DatePe);
            worksheet = entete.EnteteSRFLigne3(worksheet);
            worksheet = EcritureC(refSoc,prog,MoisPe,Annee,worksheet,NomSoc,CodeSoc);
            worksheet = EcritFooter(worksheet);      
            workbook.Worksheets.Add(worksheet);
            workbook.Save(nomFichier);
            prog.Visible = false;
            prog.Value = 0;
                Workbook book = Workbook.Load(nomFichier);
                Worksheet sheet = book.Worksheets[0];

                // traverse rows by Index
                for (int rowIndex = sheet.Cells.FirstRowIndex;
                       rowIndex <= sheet.Cells.LastRowIndex; rowIndex++)
                {
                    Row row = sheet.Cells.GetRow(rowIndex);
                    for (int colIndex = row.FirstColIndex;
                       colIndex <= row.LastColIndex; colIndex++)
                    {
                        Cell cell = row.GetCell(colIndex);
                    }
            }
                MessageBox.Show("File Created with Success");
            

        }

        public double VerifierSiNull(string SalNet)
        {
            double ValF =0.0;
            try
            {
                ValF = double.Parse(SalNet);
            }
            catch
            {
                ValF = 0.0;
            }
            return ValF;
        }

        public string RetourValNull(string Valeur)
        {
            string ValF = string.Empty;
            if (Valeur == "")
            {
                ValF = "0";
            }
            else
            {
                ValF = Valeur;
            }
            return ValF;
        }

    }
}
