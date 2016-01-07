using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelLibrary.SpreadSheet;
using System.Data;

namespace HarmonizedApp.Manager
{
    class Entete_SVR
    {
        ServiceDAO service = new ServiceDAO();
        FormatDate form = new FormatDate();

        public Worksheet EnteteSVRLigne1(Worksheet worksheet)
        {
            worksheet.Cells[0, 0] = new Cell("Summary Variance Report");
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

        public Worksheet EnteteSVRLigne4(Worksheet worksheet,string NomSoc)
        {
            worksheet.Cells[3, 0] = new Cell("Company name");
            worksheet.Cells[3, 1] = new Cell(NomSoc);
            return worksheet;
        }

        public Worksheet EnteteSVRLigne5(Worksheet worksheet,string idSoc)
        {
            worksheet.Cells[4, 0] = new Cell("Entity ID");
            worksheet.Cells[4, 1] = new Cell(idSoc);
            return worksheet;
        }

        public Worksheet EnteteSVRLigne6(Worksheet worksheet,string NomSoc)
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

        public Worksheet EnteteSVRLigne8(Worksheet worksheet,string Date1,string Date2)
        {
            string DT1 = form.LA_DATE_ADP(Date1);
            string DT2 = form.LA_DATE_ADP(Date2);
            worksheet.Cells[7, 0] = new Cell("Pay Cycle");
            worksheet.Cells[7, 1] = new Cell(DT1+"-"+DT2);
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
            worksheet.Cells[9, 0] = new Cell("");
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
            worksheet.Cells[11, 0] = new Cell("Pay code");
            worksheet.Cells[11, 1] = new Cell("Pay elements");
            worksheet.Cells[11, 2] = new Cell("Current month");
            worksheet.Cells[11, 3] = new Cell("Previous month");
            worksheet.Cells[11, 4] = new Cell("Variance");
            worksheet.Cells[11, 5] = new Cell("%");

            return worksheet;
        }

        public Worksheet EnteteSVRLigne13(Worksheet worksheet, int refSoc, int MoisPe)
        {
            int cpt = 0;
             DataTable dt = service.RecupererListeRub(refSoc);
             foreach (System.Data.DataRow rez in dt.Rows)
             {
                 string Code = rez["CodeRub"].ToString();
                 Char Cod = Char.Parse(Code.Substring(0,1));
                 string CodeF = ExtraireCode(Cod,Code);
                 string Libelles = rez["Libelles"].ToString();
                 worksheet = EcrireRubrique(cpt, worksheet, CodeF, Libelles);
                 cpt = cpt + 1;
             }
            return worksheet;
        }

        public string ExtraireCode(Char Code,string cf)
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

        public Worksheet EnteteSVRLigne14(Worksheet worksheet, int refSoc, int MoisPe,int AnneePe)
        {
            int cpt = 1;
            DataTable dt = service.ListeRubSVR(refSoc, cpt);
            foreach (System.Data.DataRow rez in dt.Rows)
            {
                string Code = rez["CodeRub"].ToString();
                worksheet = EcrireRubriqueA001(cpt, worksheet, Code, refSoc,MoisPe,AnneePe);
                cpt = cpt + 1;
            }
            return worksheet;
        }

        public Worksheet EcrireRubrique(int cpt, Worksheet worksheet, string Code, string Libelles)
        {
            worksheet.Cells[cpt + 13, 0] = new Cell(Code);
            worksheet.Cells[cpt + 13, 1] = new Cell(Libelles);
            return worksheet;
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

        public Worksheet EcrireRubriqueA001(int cpt, Worksheet worksheet, string Code,int refSoc,int MoisPe,int Annee)
        {
            int MoisPrec = VerifierMois(MoisPe);
            double valeur1 = 0.0;
            double valeur2 = 0.0;
            string CodeF = TesterChiffre(Code);
            string CodeFinal = TesterDed(CodeF);
            if (refSoc == 2)
            {
                valeur1 = service.RecupererValeur(refSoc, MoisPe, Annee, CodeFinal, "GROSSTONETA001");
                valeur2 = service.RecupererValeur(refSoc, MoisPrec, Annee, CodeFinal, "GROSSTONETA001");
            }
            if (refSoc == 5)
            {
                valeur1 = service.RecupererValeur(refSoc, MoisPe, Annee, CodeFinal, "GROSSTONETA002");
                valeur2 = service.RecupererValeur(refSoc, MoisPrec, Annee, CodeFinal, "GROSSTONETA002");
            }
            if (refSoc == 8)
            {
                valeur1 = service.RecupererValeur(refSoc, MoisPe, Annee, CodeFinal, "GROSSTONETA003");
                valeur2 = service.RecupererValeur(refSoc, MoisPrec, Annee, CodeFinal, "GROSSTONETA003");
            }
            if (refSoc == 7)
            {
                valeur1 = service.RecupererValeur(refSoc, MoisPe, Annee, CodeFinal, "GROSSTONETA004");
                valeur2 = service.RecupererValeur(refSoc, MoisPrec, Annee, CodeFinal, "GROSSTONETA004");
            }
            if (refSoc == 9)
            {
                valeur1 = service.RecupererValeur(refSoc, MoisPe, Annee, CodeFinal, "GROSSTONETA010");
                valeur2 = service.RecupererValeur(refSoc, MoisPrec, Annee, CodeFinal, "GROSSTONETA010");
            }
            if (refSoc == 10)
            {
                valeur1 = service.RecupererValeur(refSoc, MoisPe, Annee, CodeFinal, "GROSSTONETA011");
                valeur2 = service.RecupererValeur(refSoc, MoisPrec, Annee, CodeFinal, "GROSSTONETA011");
            }
            if (refSoc == 11)
            {
                valeur1 = service.RecupererValeur(refSoc, MoisPe, Annee, CodeFinal, "GROSSTONETA013");
                valeur2 = service.RecupererValeur(refSoc, MoisPrec, Annee, CodeFinal, "GROSSTONETA013");
            }
            if (refSoc == 12)
            {
                valeur1 = service.RecupererValeur(refSoc, MoisPe, Annee, CodeFinal, "GROSSTONETA015");
                valeur2 = service.RecupererValeur(refSoc, MoisPrec, Annee, CodeFinal, "GROSSTONETA015");
            }
            if (refSoc == 3)
            {
                valeur1 = service.RecupererValeur(refSoc, MoisPe, Annee, CodeFinal, "GROSSTONETA014");
                valeur2 = service.RecupererValeur(refSoc, MoisPrec, Annee, CodeFinal, "GROSSTONETA014");
            }
            if (refSoc == 4)
            {
                valeur1 = service.RecupererValeur(refSoc, MoisPe, Annee, CodeFinal, "GROSSTONETA020");
                valeur2 = service.RecupererValeur(refSoc, MoisPrec, Annee, CodeFinal, "GROSSTONETA020");
            }

            if (refSoc == 13)
            {
                valeur1 = service.RecupererValeur(refSoc, MoisPe, Annee, CodeFinal, "GROSSTONETA016");
                valeur2 = service.RecupererValeur(refSoc, MoisPrec, Annee, CodeFinal, "GROSSTONETA016");
            }

            worksheet.Cells[cpt + 12, 2] = new Cell(valeur1);
            worksheet.Cells[cpt + 12, 3] = new Cell(valeur2);
            double variance = valeur1 - valeur2;
            string pourcent = RecupererVal(valeur1, valeur2, variance);
            worksheet.Cells[cpt + 12, 4] = new Cell(variance);
            worksheet.Cells[cpt + 12, 5] = new Cell(pourcent+"%");
            return worksheet;
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
                ChaineF=Code;
            }
            return ChaineF;
        }

        public string RecupererVal(double valeur1,double valeur2,double variance)
        {
            double pourcent = 0.0;
            string ValF = string.Empty;
            if (variance == 0)
            {
                ValF = "0";
            }
            if (valeur1 == variance)
            {
                ValF = "100";
            }
            if ((variance !=0 ) && (valeur1 != variance))
            {
                pourcent = (variance / valeur2) * 100;
                ValF = pourcent.ToString("#.##");
                if (ValF == "")
                {
                    ValF = "0" + ValF;
                }
                else
                {
                    if ((ValF.Substring(0, 1) == ","))
                    {
                        ValF = "0" + ValF;
                    }
                }
            }
            if((valeur1==0)&&(valeur2==0))
            {
                 ValF = "0";
            }
            return ValF;
        }

        public int VerifierMois(int MoisPe)
        {
            int mois=0;
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

        public string verifier(double BRUT,double BRUTPREC,double Var)
        {
            double Perc=0.0;
            if (Var == 0)
            {
                Perc = 0;
            }
            if (BRUT == Var)
            {
                Perc = 1;
            }
            if ((Var > 0)&&(BRUT != Var))
            {
                Perc = (Var / BRUTPREC)*100;
            }

            return Perc.ToString("#.##");

        }

    }
}
