using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelLibrary.SpreadSheet;
using System.Data;
using System.Configuration;


namespace HarmonizedApp.Manager
{
    class Entete_SVE
    {
        ServiceDAO service = new ServiceDAO();
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
        string A016 = ConfigurationManager.AppSettings["A016"];
        string A021 = ConfigurationManager.AppSettings["A021"];


        public Worksheet EnteteSVRLigne1(Worksheet worksheet)
        {
            worksheet.Cells[0, 0] = new Cell("Summary variance report E");
            return worksheet;
        }

        public Worksheet EnteteSVRLigne2(Worksheet worksheet)
        {
            worksheet.Cells[1, 0] = new Cell("Report Date and time");
            string Year = DateTime.Today.Year.ToString();
            string Month = DateTime.Today.Month.ToString();
            string Day = DateTime.Today.Day.ToString();
            if (Month.Length == 1)
            {
                Month = "0" + Month;
            }
            if (Day.Length == 1)
            {
                Day = "0" + Day;
            }

            string DATE1 = Year + "/" + Month + "/" + Day;
            worksheet.Cells[1, 1] = new Cell(DATE1);
            return worksheet;
        }

        public Worksheet EnteteSVRLigne3(Worksheet worksheet)
        {
            worksheet.Cells[2, 0] = new Cell("Country code");
            worksheet.Cells[2, 1] = new Cell("TN");
            return worksheet;
        }

        public Worksheet EnteteSVRLigne4(Worksheet worksheet, string NomSoc)
        {
            worksheet.Cells[3, 0] = new Cell("Company name");
            worksheet.Cells[3, 1] = new Cell(NomSoc);
            return worksheet;
        }

        public Worksheet EnteteSVRLigne5(Worksheet worksheet, string idSoc)
        {
            worksheet.Cells[4, 0] = new Cell("Entity ID");
            worksheet.Cells[4, 1] = new Cell(idSoc);
            return worksheet;
        }

        public Worksheet EnteteSVRLigne6(Worksheet worksheet, string NomSoc)
        {
            worksheet.Cells[5, 0] = new Cell("Entity name");
            worksheet.Cells[5, 1] = new Cell(NomSoc);
            return worksheet;
        }

        public Worksheet EnteteSVRLigne7(Worksheet worksheet)
        {
            worksheet.Cells[6, 0] = new Cell("Currency");
            worksheet.Cells[6, 1] = new Cell("TND");
            return worksheet;
        }

        public Worksheet EnteteSVRLigne8(Worksheet worksheet, string Date1, string Date2)
        {
            string DT1 = form.LA_DATE_ADP(Date1);
            string DT2 = form.LA_DATE_ADP(Date2);
            worksheet.Cells[7, 0] = new Cell("Pay Cycle");
            worksheet.Cells[7, 1] = new Cell(DT1 + "-" + DT2);
            return worksheet;
        }

        public Worksheet EnteteSVRLigne9(Worksheet worksheet)
        {
            worksheet.Cells[8, 0] = new Cell("");
            worksheet.Cells[8, 1] = new Cell("");
            return worksheet;
        }

        public Worksheet EnteteSVRLigne10(Worksheet worksheet)
        {
            worksheet.Cells[9, 0] = new Cell("Breakdown per employee");
            worksheet.Cells[9, 1] = new Cell("");
            return worksheet;
        }

        public Worksheet EnteteSVRLigne11(Worksheet worksheet)
        {
            worksheet.Cells[10, 0] = new Cell("");
            worksheet.Cells[10, 1] = new Cell("");
            return worksheet;
        }

        public Worksheet EnteteSVRLigne12(Worksheet worksheet)
        {
            worksheet.Cells[11, 0] = new Cell("Employee ID");
            worksheet.Cells[11, 1] = new Cell("Employee Surname");
            worksheet.Cells[11, 2] = new Cell("Employee Name");
            worksheet.Cells[11, 3] = new Cell("Hire date");
            worksheet.Cells[11, 4] = new Cell("Termination date");
            worksheet.Cells[11, 5] = new Cell("Pay code");
            worksheet.Cells[11, 6] = new Cell("Pay elements");
            worksheet.Cells[11, 7] = new Cell("Current month");
            worksheet.Cells[11, 8] = new Cell("Previous month");
            worksheet.Cells[11, 9] = new Cell("Variance");
            worksheet.Cells[11, 10] = new Cell("%");

            return worksheet;
        }

        public Worksheet EnteteSVRLigne13(Worksheet worksheet, int refSoc, int MoisPe)
        {
            int cpt = 0;
            
                DataTable dt = service.RecupererListeRub(refSoc);
                foreach (System.Data.DataRow rez in dt.Rows)
                {
                    string Code = rez["CodeRub"].ToString();
                    Char Cod = Char.Parse(Code.Substring(0, 1));
                    string CodeF = ExtraireCode(Cod, Code);
                    string Libelles = rez["Libelles"].ToString();
                    worksheet = EcrireRubrique(cpt, worksheet, CodeF, Libelles);
                    cpt = cpt + 1;
                }
            return worksheet;
        }

        public string ExtraireCode(Char Code, string cf)
        {
            string Codage = string.Empty;
            bool isNumeric = Char.IsNumber(Code);  // true
            if (isNumeric)
            {
                Codage = cf.ToString().Substring(0, 4);
            }
            else
            {
                Codage = cf.ToString();
            }
            return Codage;

        }

        public Worksheet EcrireRubrique(int cpt, Worksheet worksheet, string Code, string Libelles)
        {
            worksheet.Cells[cpt + 13, 0] = new Cell(Code);
            worksheet.Cells[cpt + 13, 1] = new Cell(Libelles);
            return worksheet;
        }

        public Worksheet EnteteSVRLigne14(Worksheet worksheet,int refSoc,int MoisPe,int AnneePe)
        {
            int MoisPrec = VerifierMois(MoisPe);
            string Matricule = string.Empty;
            string CodeF = string.Empty;
            DataTable ds = null;
             int cpt = 1;

             if (refSoc == 2)
             {
                 ds = service.ListeEmployee(A001, MoisPe, AnneePe);
             }
             if (refSoc == 5)
             {
                 ds = service.ListeEmployee(A002, MoisPe, AnneePe);
             }
             if (refSoc == 8)
             {
                 ds = service.ListeEmployee(A003, MoisPe, AnneePe);
             }
             if (refSoc == 7)
             {
                 ds = service.ListeEmployee(A004, MoisPe, AnneePe);
             }
             if (refSoc == 9)
             {
                 ds = service.ListeEmployee(A010, MoisPe, AnneePe);
             }
             if (refSoc == 10)
             {
                 ds = service.ListeEmployee(A011, MoisPe, AnneePe);
             }
             if (refSoc == 11)
             {
                 ds = service.ListeEmployee(A013, MoisPe, AnneePe);
             }
             if (refSoc == 3)
             {
                 ds = service.ListeEmployee(A014, MoisPe, AnneePe);
             }
             if (refSoc == 12)
             {
                 ds = service.ListeEmployee(A015, MoisPe, AnneePe);
             }

             if (refSoc == 4)
             {
                 ds = service.ListeEmployee(A020, MoisPe, AnneePe);
             }
             if (refSoc == 13)
             {
                 ds = service.ListeEmployee(A016, MoisPe, AnneePe);
             }

             if (refSoc == 14)
             {
                 ds = service.ListeEmployee(A021, MoisPe, AnneePe);
             }

           
            foreach (System.Data.DataRow re in ds.Rows)
            {
                if ((refSoc == 11) || (refSoc == 9)||(refSoc==3)||(refSoc==4))
                {
                    Matricule = re["IdOracle"].ToString();
                }
                else
                {
                    Matricule = re["Matricule"].ToString();
                }
                string Nom = re["Nom"].ToString();
                string Prenom = re["Prenom"].ToString();
                string DateEmbauche = re["DateEmbauche"].ToString();
                string DateDepart = string.Empty;
                if (refSoc == 13)
                {
                     DateDepart = re["DateDept"].ToString();
                }
                else
                {
                     DateDepart = re["DateDepart"].ToString();
                }
               
                DataTable dt = service.ListeRubSVR(refSoc, cpt);
                foreach (System.Data.DataRow rez in dt.Rows)
                {
                    string Code = rez["CodeRub"].ToString();
                    string Libelles = rez["Libelles"].ToString();
                    CodeF = TesterChiffre(Code);
                    if (refSoc == 2)
                    {
                        TraitementDuSVE(worksheet, refSoc, MoisPe, AnneePe, CodeF, A001, Matricule,
                      cpt, Nom, Prenom, DateEmbauche, DateDepart, Libelles);
                    }
                    if (refSoc == 5)
                    {
                        TraitementDuSVE(worksheet, refSoc, MoisPe, AnneePe, CodeF, A002, Matricule,
                      cpt, Nom, Prenom, DateEmbauche, DateDepart, Libelles);
                    }
                    if (refSoc == 8)
                    {
                        TraitementDuSVE(worksheet, refSoc, MoisPe, AnneePe, CodeF, A003, Matricule,
                      cpt, Nom, Prenom, DateEmbauche, DateDepart, Libelles);
                    }
                    if (refSoc == 7)
                    {
                        TraitementDuSVE(worksheet, refSoc, MoisPe, AnneePe, CodeF, A004, Matricule,
                      cpt, Nom, Prenom, DateEmbauche, DateDepart, Libelles);
                    }
                    if (refSoc == 9)
                    {
                        TraitementDuSVE(worksheet, refSoc, MoisPe, AnneePe, CodeF, A010, Matricule,
                      cpt, Nom, Prenom, DateEmbauche, DateDepart, Libelles);
                    }
                    if (refSoc == 10)
                    {
                        TraitementDuSVE(worksheet, refSoc, MoisPe, AnneePe, CodeF, A011, Matricule,
                      cpt, Nom, Prenom, DateEmbauche, DateDepart, Libelles);
                    }
                    if (refSoc == 11)
                    {
                        TraitementDuSVE(worksheet, refSoc, MoisPe, AnneePe, CodeF, A013, Matricule,
                      cpt, Nom, Prenom, DateEmbauche, DateDepart, Libelles);
                    }
                    if (refSoc == 3)
                    {
                        TraitementDuSVE(worksheet, refSoc, MoisPe, AnneePe, CodeF, A014, Matricule,
                      cpt, Nom, Prenom, DateEmbauche, DateDepart, Libelles);
                    }
                    if (refSoc == 12)
                    {
                        TraitementDuSVE(worksheet, refSoc, MoisPe, AnneePe, CodeF, A015, Matricule,
                      cpt, Nom, Prenom, DateEmbauche, DateDepart, Libelles);
                    }
                    if (refSoc == 4)
                    {
                        TraitementDuSVE(worksheet, refSoc, MoisPe, AnneePe, CodeF, A020, Matricule,
                      cpt, Nom, Prenom, DateEmbauche, DateDepart, Libelles);
                    }
                    if (refSoc == 13)
                    {
                        TraitementDuSVE(worksheet, refSoc, MoisPe, AnneePe, CodeF, A016, Matricule,
                      cpt, Nom, Prenom, DateEmbauche, DateDepart, Libelles);
                    }
                    cpt = cpt + 1;
                }
            }
            return worksheet;
        }

        public void TraitementDuSVE(Worksheet worksheet, int refSoc,int MoisPe,int AnneePe,string Code,string Base,string Matricule,
        int cpt,string Nom,string Prenom,string DateEmbauche,string DateDepart,string Libelles)
        {
            double valeur1 = 0.0;
            double valeur2 = 0.0;
            int MoisPrec = VerifierMois(MoisPe);
            string CodeF = EffacerlaA(Code);
            string LeCode = TesterDed(Code);
            if ((refSoc == 11) || (refSoc == 9) || (refSoc == 3) || (refSoc == 4))
            {
                valeur1 = service.RecupererSVEIdOra(refSoc, MoisPe, AnneePe, LeCode, Base, Matricule);
                valeur2 = service.RecupererSVEIdOra(refSoc, MoisPrec, AnneePe, LeCode, Base, Matricule);
            }
            else
            {
                valeur1 = service.RecupererSVE(refSoc, MoisPe, AnneePe, LeCode, Base, Matricule);
                valeur2 = service.RecupererSVE(refSoc, MoisPrec, AnneePe, LeCode, Base, Matricule);
            }
           
            worksheet.Cells[cpt + 12, 0] = new Cell(Matricule);
            worksheet.Cells[cpt + 12, 1] = new Cell(Nom);
            worksheet.Cells[cpt + 12, 2] = new Cell(Prenom);

            if (DateEmbauche != "")
            {
                worksheet.Cells[cpt + 11, 3] = new Cell(form.DateAmeric(DateEmbauche));
            }
            else
            {
                worksheet.Cells[cpt + 11, 3] = new Cell(DateEmbauche);
            }
            if (DateDepart != "")
            {
                worksheet.Cells[cpt + 11, 4] = new Cell(form.DateAmeric(DateDepart));
            }
            else
            {
                worksheet.Cells[cpt + 11, 4] = new Cell(DateDepart);
            }
            worksheet.Cells[cpt + 12, 5] = new Cell(CodeF);
            worksheet.Cells[cpt + 12, 6] = new Cell(Libelles);
            worksheet.Cells[cpt + 12, 7] = new Cell(valeur1);
            worksheet.Cells[cpt + 12, 8] = new Cell(valeur2);
            double variance = valeur1 - valeur2;
            string pourcent = RecupererVal(valeur1, valeur2, variance);
            worksheet.Cells[cpt + 12, 9] = new Cell(variance);
            worksheet.Cells[cpt + 12, 10] = new Cell(pourcent + "%");
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

        public string EffacerlaA(string Code)
        {
            string Chaine = Code.Substring(0, 1);
            string ChaineF = string.Empty;
            string Test = Code.Substring(0, 1);
                if (Test == "A")
                {
                    ChaineF = Code.Substring(1, Code.Length-1);
                }
                else
                {
                    ChaineF = Code;
                }
                
            return ChaineF;
        }

        public string TesterDed(string CodeF)
        {
            string Selstatus = string.Empty;
            switch (CodeF)
            {
                case "DED":
                    Selstatus = "RUB_DEDSA";
                    break;
                case "CTR":
                    Selstatus = "RUB_COTEPY";
                    break;
                default:
                    Selstatus = CodeF;
                    break;
            }

            return Selstatus;
        }


        public string RecupererVal(double valeur1, double valeur2, double variance)
        {
            double pourcent = 0.0;
            string ValF = string.Empty;
            if (variance == 0)
            {
                ValF = "0,00";
            }
            if (valeur1 == variance)
            {
                ValF = "1,00";
            }
            if ((variance != 0) && (valeur1 != variance))
            {
                pourcent = (variance / valeur2) * 100;
                ValF = pourcent.ToString("#.##");
            }
            return ValF;
        }


        public int VerifierMois(int MoisPe)
        {
            int mois = 0;
            if (MoisPe == 1)
            {
                mois = 12;
            }
            else
            {
                mois = MoisPe - 1;
            }
            return mois;
        }

        public string verifier(double BRUT, double BRUTPREC, double Var)
        {
            double Perc = 0.0;
            if (Var == 0)
            {
                Perc = 0;
            }
            if (BRUT == Var)
            {
                Perc = 1;
            }
            if ((Var > 0) && (BRUT != Var))
            {
                Perc = (Var / BRUTPREC) * 100;
            }

            return Perc.ToString("#.##");

        }



    }
}
