using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelLibrary.SpreadSheet;
using System.Data;
using System.Configuration;

namespace HarmonizedApp.Manager
{
    class EnteteSGR
    {

        ServiceDAO service = new ServiceDAO();
        ManagerDAOSQL man = new ManagerDAOSQL();
        FormatDate form = new FormatDate();
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


        public Worksheet EnteteSVRLigne1(Worksheet worksheet)
        {
            worksheet.Cells[0, 0] = new Cell("Date du jour");
            worksheet.Cells[0, 1] = new Cell("Date de fin de période demandée");
            worksheet.Cells[0, 2] = new Cell("Code Pays société");
            worksheet.Cells[0, 3] = new Cell("Nom de la société");
            worksheet.Cells[0, 4] = new Cell("LID société");
            worksheet.Cells[0, 5] = new Cell("Nom de l'entité");
            worksheet.Cells[0, 6] = new Cell("Currency");
            worksheet.Cells[0, 7] = new Cell("Compte comptable");
            worksheet.Cells[0, 8] = new Cell("Libellé du compte");
            worksheet.Cells[0, 9] = new Cell("Code centre de cout");
            worksheet.Cells[0, 10] = new Cell("Libéllé centre de cout");
            worksheet.Cells[0, 11] = new Cell("Code Rubrique paie");
            worksheet.Cells[0, 12] = new Cell("Intitulé  Rubrique paie");
            worksheet.Cells[0, 13] = new Cell("Débit");
            worksheet.Cells[0, 14] = new Cell("Crédit");

            return worksheet;
        }

        public Worksheet EnteteOPR8Ligne1(Worksheet worksheet,string NomSoc)
        {
            worksheet.Cells[0, 0] = new Cell("Client");
            worksheet.Cells[0, 1] = new Cell(NomSoc);

            return worksheet;
        }
        public Worksheet EnteteOPR8Ligne2(Worksheet worksheet)
        {
            worksheet.Cells[1, 0] = new Cell("OPR 08");
            worksheet.Cells[1, 1] = new Cell("Comparatif Payement par Virement");

            return worksheet;
        }

        public Worksheet EnteteOPR11Ligne2(Worksheet worksheet)
        {
            worksheet.Cells[1, 0] = new Cell("OPR 11");
            worksheet.Cells[1, 1] = new Cell("OT Report");

            return worksheet;
        }

        public Worksheet EnteteOPR17Ligne2(Worksheet worksheet)
        {
            worksheet.Cells[1, 0] = new Cell("OPR 17");
            worksheet.Cells[1, 1] = new Cell("Payroll summary");

            return worksheet;
        }

        public Worksheet EnteteOPR15Ligne2(Worksheet worksheet)
        {
            worksheet.Cells[1, 0] = new Cell("OPR 15");
            worksheet.Cells[1, 1] = new Cell("FINANCIAL REPORT");

            return worksheet;
        }


        public Worksheet EnteteOPR8Ligne3(Worksheet worksheet,int Mois)
        {
            worksheet.Cells[3, 0] = new Cell("Month");
            worksheet.Cells[3, 1] = new Cell(Mois);

            return worksheet;
        }

        public Worksheet EnteteOPR8Ligne4(Worksheet worksheet, int Annee)
        {
            worksheet.Cells[4, 0] = new Cell("YEAR");
            worksheet.Cells[4, 1] = new Cell(Annee);

            return worksheet;
        }

        public Worksheet EnteteOPR8Ligne5(Worksheet worksheet)
        {
            worksheet.Cells[5, 0] = new Cell("local ID");
            worksheet.Cells[5, 1] = new Cell("Name");
            worksheet.Cells[5, 2] = new Cell("First Name");
            worksheet.Cells[5, 3] = new Cell("Account actual month");
            worksheet.Cells[5, 4] = new Cell("Net to pay Actual month");
            worksheet.Cells[5, 5] = new Cell("Basic salary actual month");
            worksheet.Cells[5, 6] = new Cell("Account previous month");
            worksheet.Cells[5, 7] = new Cell("Net to pay  Previous month");
            worksheet.Cells[5, 8] = new Cell("Basic salary  Previous month");
            worksheet.Cells[5, 9] = new Cell("Actual month");
            worksheet.Cells[5, 10] = new Cell("Comment");

            return worksheet;
        }

        public Worksheet EnteteOPR11Ligne5(Worksheet worksheet)
        {
            worksheet.Cells[5, 0] = new Cell("ID Groupe");
            worksheet.Cells[5, 1] = new Cell("local ID");
            worksheet.Cells[5, 2] = new Cell("Name");
            worksheet.Cells[5, 3] = new Cell("First Name");
            worksheet.Cells[5, 4] = new Cell("Employement Title");
            worksheet.Cells[5, 5] = new Cell("Element code");
            worksheet.Cells[5, 6] = new Cell("Description");
            worksheet.Cells[5, 7] = new Cell("Nbre Hour");
            worksheet.Cells[5, 8] = new Cell("Amount");
            worksheet.Cells[5, 9] = new Cell("Cost center");
            worksheet.Cells[5, 10] = new Cell("Month Pay");

            return worksheet;
        }

        public Worksheet EnteteOPR17Ligne5(Worksheet worksheet)
        {
            worksheet.Cells[5, 0] = new Cell("Code");
            worksheet.Cells[5, 1] = new Cell("Rubriques");
            worksheet.Cells[5, 2] = new Cell("Gains");
            worksheet.Cells[5, 3] = new Cell("Cumuls Annuel");

            return worksheet;
        }

        public Worksheet EnteteOPR15Ligne5(Worksheet worksheet)
        {
            worksheet.Cells[5, 0] = new Cell("PK");
            worksheet.Cells[5, 1] = new Cell("Account");
            worksheet.Cells[5, 2] = new Cell("Account short text");
            worksheet.Cells[5, 3] = new Cell("Tx");
            worksheet.Cells[5, 4] = new Cell("Cost Ctr");
            worksheet.Cells[5, 5] = new Cell("Profit Ctr");
            worksheet.Cells[5, 6] = new Cell("Order");
            worksheet.Cells[5, 7] = new Cell("Amount");
            worksheet.Cells[5, 8] = new Cell(" Period");

            return worksheet;
        }


        public Worksheet EnteteOPR8Ligne6(Worksheet worksheet,int Mois1,int AnneeP1,int MoisPrec,int AnneePrec)
        {
            int cpt = 1;
            DataTable dt = service.ListeOPR8(Mois1, AnneeP1);
            foreach (System.Data.DataRow rez in dt.Rows)
            {

                worksheet.Cells[5 + cpt, 0] = new Cell(rez["Matricule"].ToString());
                worksheet.Cells[5 + cpt, 1] = new Cell(rez["Nom"].ToString());
                worksheet.Cells[5 + cpt, 2] = new Cell(rez["Prenom"].ToString());
                worksheet.Cells[5 + cpt, 3] = new Cell(rez["RIB"].ToString());
                worksheet.Cells[5 + cpt, 4] = new Cell(rez["NETPAIE"].ToString());
                worksheet.Cells[5 + cpt, 5] = new Cell(rez["A1000"].ToString());
                worksheet.Cells[5 + cpt, 6] = new Cell(service.RIBBGMOISPREC(rez["Matricule"].ToString(), MoisPrec, AnneePrec));
                worksheet.Cells[5 + cpt, 7] = new Cell(service.NETPAIEBGMOISPREC(rez["Matricule"].ToString(), MoisPrec, AnneePrec));
                worksheet.Cells[5 + cpt, 8] = new Cell(service.SALBASEBGMOISPREC(rez["Matricule"].ToString(), MoisPrec, AnneePrec));
                worksheet.Cells[5 + cpt, 9] = new Cell(Mois1);
                cpt++;
            }
            return worksheet;
        }

        public Worksheet EnteteOPR11Ligne6(Worksheet worksheet, int Mois1, int AnneeP1)
        {
            int cpt = 1;
            DataTable dt = service.ListeOPR8(Mois1, AnneeP1);
            foreach (System.Data.DataRow rez in dt.Rows)
            {
                string Matricule = rez["Matricule"].ToString();
                string Nom = rez["Nom"].ToString();
                string Prenom = rez["Prenom"].ToString();
                string CodeDept = rez["CodeDept"].ToString();
                string Intitule = rez["IntituleDept"].ToString();
                string MoisPe = rez["MoisPe"].ToString();
                string Emploi = rez["Emploi"].ToString();

                DataTable dff = service.Rub11();

                foreach (System.Data.DataRow ret in dff.Rows)
                {
                    string CodeRub = ret["CodeRub"].ToString();
                    string Lib = ret["LibellesRubrique"].ToString();
                    string Test = CodeRub.Substring(0, 1);
                    if (Test == ("C") || (Test == "H"))
                    {
                        worksheet.Cells[5 + cpt, 0] = new Cell(Matricule);
                        worksheet.Cells[5 + cpt, 1] = new Cell(Matricule);
                        worksheet.Cells[5 + cpt, 2] = new Cell(Nom);
                        worksheet.Cells[5 + cpt, 3] = new Cell(Prenom);
                        worksheet.Cells[5 + cpt, 4] = new Cell(Emploi);
                        worksheet.Cells[5 + cpt, 5] = new Cell(CodeRub);
                        worksheet.Cells[5 + cpt, 6] = new Cell(Lib);
                        worksheet.Cells[5 + cpt, 7] = new Cell(double.Parse(service.ValRubA016(Matricule, Mois1, AnneeP1, CodeRub)));
                        worksheet.Cells[5 + cpt, 8] = new Cell(0);
                        worksheet.Cells[5 + cpt, 9] = new Cell(CodeDept);
                        worksheet.Cells[5 + cpt, 10] = new Cell(Intitule);
                        worksheet.Cells[5 + cpt, 11] = new Cell(MoisPe);
                        cpt++;
                    }
                    else
                    {
                        worksheet.Cells[5 + cpt, 12] = new Cell(CodeRub);
                        worksheet.Cells[5 + cpt, 13] = new Cell(Lib);
                        worksheet.Cells[5 + cpt, 14] = new Cell(double.Parse(service.ValRubA016(Matricule, Mois1, AnneeP1,"A"+CodeRub)));
                    }
                }
            }
            return worksheet;
        }

        public Worksheet EnteteOPR17Ligne6(Worksheet worksheet, int Mois1, int AnneeP1)
        {
            int cpt = 1;
            double ValRub = 0.0;
            double SumValRub = 0.0;
            DataTable dt = service.ListeRubSVR(13, cpt);
            foreach (System.Data.DataRow rez in dt.Rows)
            {
                string Code = rez["CodeRub"].ToString();
                string Libelles = rez["Libelles"].ToString();
                worksheet.Cells[5 + cpt, 0] = new Cell(Code);
                worksheet.Cells[5 + cpt, 1] = new Cell(Libelles);
                string TestCode = Code.Substring(0, 1);
                if ((TestCode == "C") || (TestCode == "H") || (TestCode == "S") || (TestCode == "B") || (TestCode == "R") || (TestCode == "N") || (TestCode == "T"))
                {
                    ValRub = service.RecupererValeur(13, Mois1, AnneeP1, Code, "GROSSTONETA016");
                }
                else
                {
                    ValRub = service.RecupererValeur(13, Mois1, AnneeP1, "A"+Code, "GROSSTONETA016");
                }
                worksheet.Cells[5 + cpt, 2] = new Cell(ValRub.ToString("#.###"));
                if ((TestCode == "C") || (TestCode == "H") || (TestCode == "S") || (TestCode == "B") || (TestCode == "R") || (TestCode == "N") || (TestCode == "T"))
                {
                    SumValRub = service.RecupererSumRubrique(13, Mois1, AnneeP1, Code, "GROSSTONETA016");
                }
                else
                {
                    string cc = "A"+Code;
                    SumValRub = service.RecupererSumRubrique(13, Mois1, AnneeP1, cc, "GROSSTONETA016");
                }
                 
                worksheet.Cells[5 + cpt, 3] = new Cell(SumValRub.ToString("#.###"));
                cpt++;
            }
            return worksheet;
        }

        public Worksheet EnteteOPR15Ligne6(Worksheet worksheet, int Mois1, int AnneeP1)
        {
            int cpt = 1;
            double ValRub = 0.0;
            DataTable dt = service.ListeRubSVR(13, cpt);
            foreach (System.Data.DataRow rez in dt.Rows)
            {
                string Code = rez["CodeRub"].ToString();
                string Libelles = rez["Libelles"].ToString();
                worksheet.Cells[5 + cpt, 0] = new Cell(Code);
                worksheet.Cells[5 + cpt, 1] = new Cell(Libelles);
               
                cpt++;
            }
            return worksheet;
        }

        public string NomTab(int refSoc)
        {
            string NomDeTab = string.Empty;
            if (refSoc == 2)
            {
                NomDeTab = A001;
            }
            if (refSoc == 5)
            {
                NomDeTab = A002;
            }
            if (refSoc == 8)
            {
                NomDeTab = A003;
            }
            if (refSoc == 7)
            {
                NomDeTab = A004;
            }
            if (refSoc == 9)
            {
                NomDeTab = A010;
            }
            if (refSoc == 10)
            {
                NomDeTab = A011;
            }
            if (refSoc == 11)
            {
                NomDeTab = A013;
            }
            if (refSoc == 3)
            {
                NomDeTab = A014;
            }
            if (refSoc == 12)
            {
                NomDeTab = A015;
            }
            if (refSoc == 4)
            {
                NomDeTab = A020;
            }

            return NomDeTab;

        }

        public Worksheet EnteteSVRLigne2(Worksheet worksheet, string Date1, string Date2, string DateJour,
        string NomSoc, string CodeSoc, string TypeTri, int refSoc, int Mois, int Annee)
        {
            int cpt = 0;
            string TabSoc = NomTab(refSoc);
            if (TypeTri == "Dept")
            {
                string CodeDept = "CodeDept";
                string IntituleDept = "IntituleDept";
                TraitementEcriture(CodeDept, IntituleDept, TabSoc, Mois, Annee, worksheet, DateJour, Date2, NomSoc, CodeSoc, cpt, refSoc,CodeDept);

            }
            if (TypeTri == "Sce")
            {
                string CodeSce = "CodeSce";
                string IntituleSce = "IntituleSce";
                TraitementEcriture(CodeSce,IntituleSce, TabSoc, Mois, Annee, worksheet, DateJour, Date2, NomSoc, CodeSoc, cpt, refSoc,CodeSce);

            }

            if (TypeTri == "CentreC")
            {
                string CodeSce = "CentreCout";
                string IntituleSce = "CentreCout";
                string CentreCout = "CentreCout";
                TraitementEcriture(CodeSce, IntituleSce, TabSoc, Mois, Annee, worksheet, DateJour, Date2, NomSoc, CodeSoc, cpt, refSoc, CentreCout);

            }

            return worksheet;
        }

        public List<DevinerSens> VoirSens(string Sens, double Valeur)
        {
          //  double ValF = 0.0;
            List<DevinerSens> devin = null;
            DevinerSens Senser = null;
            if ((Sens == "D") && (Valeur < 0))
            {
                 devin = new List<DevinerSens>();
                Senser = new DevinerSens();
                Senser.LeSens = "C";
                Senser.ValeurF = Valeur * (-1);
                devin.Add(Senser);
            }
            if ((Sens == "C") && (Valeur < 0))
            {
                devin = new List<DevinerSens>();
                Senser = new DevinerSens();
                Senser.LeSens = "D";
                Senser.ValeurF = Valeur * (-1);
                devin.Add(Senser);
            }
            return devin;
        }

        public void TraitementEcriture(string CodeSce,string IntituleSce,string TabSoc,int Mois,int Annee,Worksheet worksheet,
           string DateJour, string Date2, string NomSoc, string CodeSoc, int cpt, int refSoc,string CentreSceDept)
        {
            DataTable dt = service.ListeSceSGL(TabSoc, Mois, Annee, CodeSce, IntituleSce);
            foreach (System.Data.DataRow rez in dt.Rows)
            {
                string CodeSSS = rez[CodeSce].ToString();
                string Intitule = rez[IntituleSce].ToString();
                DataTable fg = service.ListeRubCptSGL(refSoc);
                foreach (System.Data.DataRow fy in fg.Rows)
                {
                    string CodeRub = fy["CodeRub"].ToString();
                    string LibRub = fy["LibCodeRub"].ToString();
                    string CodeCpt = fy["CodeCpt"].ToString();
                    string Libelles = fy["Libelles"].ToString();
                    string Sens = fy["Sens"].ToString();
                    string CodeRubrik = TesterChiffre(CodeRub);
                    double Valeur = service.ValeurCodeGLFILE(CodeRubrik, TabSoc, Mois, Annee, CodeSSS, CentreSceDept);

                        if (CodeRub == "COTEPY")
                        {
                            Valeur = DecTotRubCoty(NomTab(refSoc), Mois, Annee, CodeSce, CodeSSS,refSoc);
                        }
                        if (CodeRub == "ASSGEPY")
                        {
                            Valeur = AssuranceGrp(NomTab(refSoc), Mois, Annee, CodeSce, CodeSSS, refSoc);
                        }
                        string CodeFF = VerifierVal(refSoc,CodeCpt,CodeRub);
                        if (Sens == "D")
                        {
                            if (Valeur != 0)
                            {                              
                                    TraitementSGL(worksheet, DateJour, Date2, NomSoc, CodeSoc, cpt);
                                    Repetition(worksheet, cpt, CodeCpt, LibRub, Libelles, CodeSSS, Intitule, CodeFF);
                                    VerifierLaSociete(refSoc, CodeFF, worksheet, cpt, Valeur, Sens, CodeCpt);
                               
                                if (Valeur > 0)
                                {
                                    service.insertDATAGLFILE(CodeFF, CodeCpt, Libelles, CodeSSS, Intitule, Sens, Valeur.ToString("#.###"), refSoc);
                                }
                                else
                                {
                                    double ValF = (Valeur * (-1));
                                    service.insertDATAGLFILE(CodeFF, CodeCpt, Libelles, CodeSSS, Intitule, Sens, ValF.ToString("#.###"), refSoc);
                                }
                                cpt = cpt + 1;
                            }

                        }
                        else
                        {
                            if (Valeur != 0)
                            {
                                    TraitementSGL(worksheet, DateJour, Date2, NomSoc, CodeSoc, cpt);
                                    Repetition(worksheet, cpt, CodeCpt, LibRub, Libelles, CodeSSS, Intitule, CodeFF);
                                    VerifierLaSociete(refSoc, CodeFF, worksheet, cpt, Valeur, Sens, CodeCpt);
                                if (Valeur > 0)
                                {
                                    service.insertDATAGLFILE(CodeFF, CodeCpt, Libelles, CodeSSS, Intitule, Sens, Valeur.ToString("#.###"), refSoc);
                                }
                                else
                                {
                                    double ValF = (Valeur * (-1));
                                    service.insertDATAGLFILE(CodeFF, CodeCpt, Libelles, CodeSSS, Intitule, Sens, ValF.ToString("#.###"), refSoc);
                                }
                                cpt = cpt + 1;
                            }

                        }

                    }
                }
            }

        public string VerifierVal(int RefSoc,string CodeCpt,string CodeRub)
        {
            string CodeF=string.Empty;
            if ((RefSoc == 2)&&(CodeCpt=="453120"))
            {
                CodeF = "623099";
            }
                else
                {
                    CodeF = CodeRub;
                }
            return CodeF;
        }

        public void VerifierLaSociete(int refSoc,string CodeRub,Worksheet worksheet,int cpt,double Valeur,string Sens,string CodeCpt)
        {
           
            if ((CodeRub == "6320") || (CodeRub == "6330"))
            {
                TFPFOP(worksheet, cpt, Valeur, Sens);
            }
            else
            {
                    TesterSens(refSoc, worksheet, cpt, Valeur, CodeRub);
            }
            if ((refSoc == 8))
            {
                if ((CodeRub == "6210")||(CodeRub=="62109")||(CodeRub=="6120"))
                {
                    TFPFOP(worksheet, cpt, Valeur, Sens);
                }
            }
            if ((refSoc == 7))
            {
                if((CodeRub == "62109") || (CodeRub == "6120"))
                {
                    TFPFOP(worksheet, cpt, Valeur, Sens);
                }
            }
            
        }

        public void TFPFOP(Worksheet worksheet, int cpt, double Val,string Sens)
        {
            if (Sens == "D")
            {
                worksheet.Cells[cpt + 1, 13] = new Cell(Val);
                worksheet.Cells[cpt + 1, 14] = new Cell(0);
            }
            else
            {
                worksheet.Cells[cpt + 1, 13] = new Cell(0);
                worksheet.Cells[cpt + 1, 14] = new Cell(Val);
            }
        }

        public void TesterSens(int refSoc, Worksheet worksheet,int cpt,double Valeur,string CodeRub)
        {
            int Gain = 0;
            int TypRub = 0;
            int taille = CodeRub.Length;
            if (taille >= 5)
            {
                if ((CodeRub == "COTEPY") || (CodeRub == "ASSGEPY")||(CodeRub=="NETPAIE")||(CodeRub=="623099"))
                {
                    worksheet.Cells[cpt + 1, 13] = new Cell(0);
                    worksheet.Cells[cpt + 1, 14] = new Cell(Valeur);
                }
                else
                {
                    worksheet.Cells[cpt + 1, 11] = new Cell(CodeRub.Substring(0,4));
                    worksheet.Cells[cpt + 1, 13] = new Cell(Valeur);
                    worksheet.Cells[cpt + 1, 14] = new Cell(0);
                }
            }
            else
            {
                DataTable fg = man.TestSens(CodeRub, refSoc);
                foreach (System.Data.DataRow fy in fg.Rows)
                {
                    TypRub = int.Parse(fy["TypeRubrique"].ToString());
                    Gain = int.Parse(fy["Gain"].ToString());
                }
                if ((TypRub == 1))
                {
                    if ((Gain == 1))
                    {
                        if (Valeur > 0)
                        {
                            worksheet.Cells[cpt + 1, 13] = new Cell(Valeur);
                            worksheet.Cells[cpt + 1, 14] = new Cell(0);
                        }
                        else
                        {
                            worksheet.Cells[cpt + 1, 13] = new Cell(0);
                            worksheet.Cells[cpt + 1, 14] = new Cell(Valeur*(-1));
                        }
                    }
                    else
                    {
                      
                        if (Valeur > 0)
                        {
                            worksheet.Cells[cpt + 1, 13] = new Cell(0);
                            worksheet.Cells[cpt + 1, 14] = new Cell(Valeur);
                        }
                        else
                        {
                            worksheet.Cells[cpt + 1, 13] = new Cell(Valeur * (-1));
                            worksheet.Cells[cpt + 1, 14] = new Cell(0);
                        }
                        if (refSoc == 9)
                        {
                            if ((CodeRub == "4174") || (CodeRub == "4176"))
                            {
                                worksheet.Cells[cpt + 1, 13] = new Cell(0);
                                worksheet.Cells[cpt + 1, 14] = new Cell(Valeur*(-1));
                            }

                        }
                    }
                }
                if (TypRub == 3)
                {
                    if (Gain == 1)
                    {
                        if (Valeur > 0)
                        {
                            worksheet.Cells[cpt + 1, 13] = new Cell(Valeur);
                            worksheet.Cells[cpt + 1, 14] = new Cell(0);
                        }
                        else
                        {
                            worksheet.Cells[cpt + 1, 13] = new Cell(0);
                            worksheet.Cells[cpt + 1, 14] = new Cell(Valeur * (-1));
                        }
                    }
                    else
                    {
                        if (Valeur > 0)
                        {
                            worksheet.Cells[cpt + 1, 13] = new Cell(0);
                            worksheet.Cells[cpt + 1, 14] = new Cell(Valeur);
                        }
                        else
                        {
                            worksheet.Cells[cpt + 1, 13] = new Cell(Valeur * (-1));
                            worksheet.Cells[cpt + 1, 14] = new Cell(0);
                        }
                    }
                }
                if (TypRub == 2)
                {
                    if ((CodeRub == "6150")||(CodeRub == "6222"))
                    {
                        worksheet.Cells[cpt + 1, 13] = new Cell(Valeur);
                        worksheet.Cells[cpt + 1, 14] = new Cell(0);
                    }
                    else
                    {
                        worksheet.Cells[cpt + 1, 13] = new Cell(0);
                        worksheet.Cells[cpt + 1, 14] = new Cell(Valeur);
                    }
                }
                if (refSoc == 5)
                {
                    if ((CodeRub == "CL24") || (CodeRub == "CL25"))
                    {
                        worksheet.Cells[cpt + 1, 13] = new Cell(Valeur);
                        worksheet.Cells[cpt + 1, 14] = new Cell(0);
                    }
                }
            }
           
           
        }

        public void Repetition(Worksheet worksheet, int cpt, string CodeCpt,string LibRub, string Libelles, string CodeSSS, string Intitule, string CodeRub)
        {
            worksheet.Cells[cpt + 1, 7] = new Cell(CodeCpt);
            worksheet.Cells[cpt + 1, 8] = new Cell(Libelles);
            worksheet.Cells[cpt + 1, 9] = new Cell(CodeSSS);
            worksheet.Cells[cpt + 1, 10] = new Cell(Intitule);
            worksheet.Cells[cpt + 1, 11] = new Cell(CodeRub);
            worksheet.Cells[cpt + 1, 12] = new Cell(LibRub);
        }

        public double DecTotRubCoty(string NomTabb,int Mois,int Annee,string SceDept,string CodeSce,int RefSoc)
        {
            double val = 0.0;
            if (RefSoc == 7)
            {
                double A6110 = service.CnssEmplyeurF("A61109", NomTabb, Mois, Annee, SceDept, CodeSce);
                double A6120 = service.CnssEmplyeurF("A6120", NomTabb, Mois, Annee, SceDept, CodeSce);
                double A6150 = service.CnssEmplyeurF("A6150", NomTabb, Mois, Annee, SceDept, CodeSce);
                val = A6110 + A6120 + A6150;
            }
            else
            {
                double A6110 = service.CnssEmplyeurF("A61109", NomTabb, Mois, Annee, SceDept, CodeSce);
                double A6120 = service.CnssEmplyeurF("A61209", NomTabb, Mois, Annee, SceDept, CodeSce);
                double A6150 = service.CnssEmplyeurF("A6150", NomTabb, Mois, Annee, SceDept, CodeSce);
                val = A6110 + A6120 + A6150;
            }
            return val;
        }

        public string AssGrp1(int RefSoc)
        {
            string Ass1 = string.Empty;
            if (RefSoc == 3)
            {
                Ass1 = "A711899";
            }
            if (RefSoc == 4)
            {
                Ass1 = "A71159";
            }
            if (RefSoc == 2)
            {
                Ass1 = "A62109";
            }
            if (RefSoc == 7)
            {
                Ass1 = "A621099";
            }
            if (RefSoc == 5)
            {
                Ass1 = "A71159";
            }
            if (RefSoc == 11)
            {
                Ass1 = "A71159";
            }
            return Ass1;
        }

        public string AssGrp2(int RefSoc)
        {
            string Ass1 = string.Empty;
            if (RefSoc == 3)
            {
                Ass1 = "A711599";
            }
            if (RefSoc == 4)
            {
                Ass1 = "A62019";
            }
            if (RefSoc == 2)
            {
                Ass1="A62309";
            }
           
            return Ass1;
        }

        public double AssuranceGrp(string NomTabb, int Mois, int Annee, string SceDept, string CodeSce,int RefSoc)
        {
            double val = 0.0;
                double A62109 = service.CnssEmplyeurF(AssGrp1(RefSoc), NomTabb, Mois, Annee, SceDept, CodeSce);
                double A62309 = service.CnssEmplyeurF(AssGrp2(RefSoc), NomTabb, Mois, Annee, SceDept, CodeSce);
                val = A62109 + A62309;

            return val;
        }

        public string TesterChiffre(string Code)
        {
            string Chaine = Code.Substring(0, 1);
            string ChaineF = string.Empty;
            try
            {
                int Chiffre = int.Parse(Code);
                ChaineF = "A" + Code;
            }
            catch
            {
                ChaineF = Code;
            }
            return ChaineF;
        }

        public void TraitementSGL(Worksheet worksheet, string DateJour, string Date2, string NomSoc, string CodeSoc, int cpt)
        {
            worksheet.Cells[cpt + 1, 0] = new Cell(DateJour);
            worksheet.Cells[cpt + 1, 1] = new Cell(Date2);
            worksheet.Cells[cpt + 1, 2] = new Cell("TN");
            worksheet.Cells[cpt + 1, 3] = new Cell(NomSoc);
            worksheet.Cells[cpt + 1, 4] = new Cell(CodeSoc);
            worksheet.Cells[cpt + 1, 5] = new Cell(NomSoc);
            worksheet.Cells[cpt + 1, 6] = new Cell("TND");
        }

    }
}
