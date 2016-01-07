using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelLibrary.SpreadSheet;
using System.Data;
using System.Windows.Forms;

namespace HarmonizedApp.Manager
{
    class ContainSEP
    {

        ServiceDAO service = new ServiceDAO();
        FormatDate form = new FormatDate();

        public Worksheet EnteteSGNLigne1(Worksheet worksheet)
        {
            worksheet.Cells[0, 0] = new Cell("Payment Report");
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

        public Worksheet EnteteSVRLigne5(Worksheet worksheet,string CodeSoc)
        {
            worksheet.Cells[4, 0] = new Cell("Entity ID");
            worksheet.Cells[4, 1] = new Cell(CodeSoc);
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
            worksheet.Cells[8, 0] = new Cell("breakdown per employee");
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
            worksheet.Cells[10, 0] = new Cell("Matricule");
            worksheet.Cells[10, 1] = new Cell("Nom");
            worksheet.Cells[10, 2] = new Cell("Prénom");
            worksheet.Cells[10, 3] = new Cell("Date d'embauche société");
            worksheet.Cells[10, 4] = new Cell("Date de départ société");
            worksheet.Cells[10, 5] = new Cell("Mode Paiement");
            worksheet.Cells[10, 6] = new Cell("Information Bancaire");
            worksheet.Cells[10, 7] = new Cell("Net à Payer");

            return worksheet;
        }

        public Worksheet EnteteSVRLigne12(Worksheet worksheet, int refSoc, int MoisPe,int annee)
        {
            int cpt = 0;
            DataTable dt = service.RecupererDATASEP(refSoc, MoisPe,annee);
            foreach (System.Data.DataRow rez in dt.Rows)
            {
               
                string Mat = rez["Matricule"].ToString();
                string Nom = rez["Nom"].ToString();
                string Prenom = rez["Prenom"].ToString();
                string DateEmbauche = rez["DateEmbauche"].ToString();
                string DateDepart = rez["DateDepart"].ToString();
                string ModePaie = rez["ModePe"].ToString();
                string InfoBank = rez["InfoBank"].ToString();
                string NetPaie = rez["NetPaie"].ToString();
                worksheet = EcrireSEP(cpt, worksheet, Mat, Nom, Prenom, DateEmbauche, DateDepart,ModePaie, InfoBank, NetPaie);
                cpt = cpt + 1;
            }

            return worksheet;
        }

        public Worksheet EcrireSEP(int cpt, Worksheet worksheet, string Mat, string Nom, string Prenom, string DateEmbauche, string DateDepart,
           string ModePaie, string InfoBank, string NetPaie)
        {
            worksheet.Cells[cpt + 11, 0] = new Cell(Mat);
            worksheet.Cells[cpt + 11, 1] = new Cell(Nom);
            worksheet.Cells[cpt + 11, 2] = new Cell(Prenom);

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

            worksheet.Cells[cpt + 11, 5] = new Cell(ModePaie);
            worksheet.Cells[cpt + 11, 6] = new Cell(InfoBank);
            worksheet.Cells[cpt + 11, 7] = new Cell(NetPaie);
            return worksheet;
        }

    }
}
