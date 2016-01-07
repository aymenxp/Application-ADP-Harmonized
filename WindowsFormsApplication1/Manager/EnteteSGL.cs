using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelLibrary.SpreadSheet;
using System.Data;
using System.Configuration;

namespace HarmonizedApp.Manager
{
    class EnteteSGL
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


        public Worksheet EnteteSVRLigne1(Worksheet worksheet)
        {
            worksheet.Cells[0, 0] = new Cell("Report date");
            worksheet.Cells[0, 1] = new Cell("Pay cycle");
            worksheet.Cells[0, 2] = new Cell("Country code");
            worksheet.Cells[0, 3] = new Cell("Company name");
            worksheet.Cells[0, 4] = new Cell("Entity ID");
            worksheet.Cells[0, 5] = new Cell("Entity name");
            worksheet.Cells[0, 6] = new Cell("Currency");
            worksheet.Cells[0, 7] = new Cell("GL account");
            worksheet.Cells[0, 8] = new Cell("GL description");
            worksheet.Cells[0, 9] = new Cell("Cost center code");
            worksheet.Cells[0, 10] = new Cell("Cost center description");
            worksheet.Cells[0, 11] = new Cell("Amount debit");
            worksheet.Cells[0, 12] = new Cell("Amount credit");

            return worksheet;
        }

        public string NomTab(int refSoc)
        {
            string NomDeTab=string.Empty;
            if(refSoc==2)
            {
                NomDeTab=A001;
            }
             if(refSoc==5)
            {
                NomDeTab=A002;
            }
             if(refSoc==8)
            {
                NomDeTab=A003;
            }
             if(refSoc==7)
            {
                NomDeTab=A004;
            }
              if(refSoc==9)
            {
                NomDeTab=A010;
            }
             if(refSoc==10)
            {
                NomDeTab=A011;
            }
              if(refSoc==11)
            {
                NomDeTab=A013;
            }
              if(refSoc==3)
            {
                NomDeTab=A014;
            }
              if(refSoc==12)
            {
                NomDeTab=A015;
            }
             if(refSoc==4)
            {
                NomDeTab=A020;
            }

            return NomDeTab;

        }
  
        public Worksheet EnteteSVRLigne2(Worksheet worksheet, string Date1, string Date2, string DateJour,
        string NomSoc, string CodeSoc, string TypeTri, int refSoc, int Mois, int Annee)
        {
            int cpt = 0;
            int compteur = 1;
            DataTable dz = service.ListGlAccount();
             foreach (System.Data.DataRow re in dz.Rows)
             {
                 string COSTCENTER = re["COSTCENTER"].ToString();
                 DataTable de = service.CodeCptF(COSTCENTER);
                 foreach (System.Data.DataRow fd in de.Rows)
                 {
                     string GLACCOUNT = fd["GLACCOUNT"].ToString();
                     string SENS = fd["SENS"].ToString();
                     DataTable bc = service.ListDeGlFileData(COSTCENTER,GLACCOUNT,SENS);                 
                     foreach (System.Data.DataRow qed in bc.Rows)
                     {
                         if (bc.Rows.Count == 1)
                         {
                            
                            EcrireMethode(qed,worksheet, cpt, DateJour, Date2, NomSoc, CodeSoc, GLACCOUNT,COSTCENTER);
                            cpt = cpt + 1;
                         }
                         else
                         {
                             if ((bc.Rows.Count == compteur))
                             {
                                 EcrireMethode(qed, worksheet, cpt, DateJour, Date2, NomSoc, CodeSoc, GLACCOUNT, COSTCENTER);
                                 compteur = 0;
                                 cpt = cpt + 1;
                             }                            
                             compteur = compteur + 1;
                         }
                        
                     }
                
                 }
                 
             }
          
           return worksheet;

        }

        public void EcrireMethode(DataRow qed, Worksheet worksheet, int cpt, string DateJour, string Date2, string NomSoc, string CodeSoc, string GLACCOUNT, string COSTCENTER)
        {
            string GLACC = qed["GLACCOUNT"].ToString();
            string GLDESCRIPTION = qed["GLDESCRIPTION"].ToString();
            string COSTCENTER1 = qed["COSTCENTER"].ToString();
            string DESCRIPTION = qed["DESCRIPTION"].ToString();
            string SENS = qed["SENS"].ToString();
            double VALEUR = service.ValeurFinalSens(GLACCOUNT, COSTCENTER, SENS);
            InsererData(worksheet, cpt, DateJour, Date2, NomSoc, CodeSoc, GLACC, GLDESCRIPTION, COSTCENTER, DESCRIPTION, SENS, VALEUR);
            
        }

        public void InsererData(Worksheet worksheet, int cpt, string DateJour, string Date2, string NomSoc, string CodeSoc, string GLACCOUNT,
        string GLDESCRIPTION, string COSTCENTER, string DESCRIPTION,string Sens,double Val)
        {
                 
                    worksheet.Cells[cpt + 1, 0] = new Cell(DateJour);
                    worksheet.Cells[cpt + 1, 1] = new Cell(Date2);
                    worksheet.Cells[cpt + 1, 2] = new Cell("TN");
                    worksheet.Cells[cpt + 1, 3] = new Cell(NomSoc);
                    worksheet.Cells[cpt + 1, 4] = new Cell(CodeSoc);
                    worksheet.Cells[cpt + 1, 5] = new Cell(NomSoc);
                    worksheet.Cells[cpt + 1, 6] = new Cell("TND");
                    worksheet.Cells[cpt + 1, 7] = new Cell(GLACCOUNT);
                    worksheet.Cells[cpt + 1, 8] = new Cell(GLDESCRIPTION);
                    worksheet.Cells[cpt + 1, 9] = new Cell(COSTCENTER);
                    worksheet.Cells[cpt + 1, 10] = new Cell(DESCRIPTION);
                    if (Sens == "C")
                    {
                            worksheet.Cells[cpt + 1, 11] = new Cell(0);
                            worksheet.Cells[cpt + 1, 12] = new Cell(Val);
                 
                    }
                    else
                    {
                        worksheet.Cells[cpt + 1, 11] = new Cell(Val);
                        worksheet.Cells[cpt + 1, 12] = new Cell(0);
                    }
                    if ((GLACCOUNT == "453130") || (GLACCOUNT == "457000"))
                    {
                        if (Sens == "C")
                        {
                            worksheet.Cells[cpt + 1, 11] = new Cell(Val);
                            worksheet.Cells[cpt + 1, 12] = new Cell(0);
                        }
                        else
                        {
                            if (Val > 0)
                            {
                                worksheet.Cells[cpt + 1, 11] = new Cell(0);
                                worksheet.Cells[cpt + 1, 12] = new Cell(Val);
                            }
                            else
                            {
                                double VF = (Val * (-1));
                                worksheet.Cells[cpt + 1, 11] = new Cell(0);
                                worksheet.Cells[cpt + 1, 12] = new Cell(VF);
                            }
                           
                        }
                    }
             if ((GLACCOUNT == "647100") || (GLACCOUNT == "647200")||(GLACCOUNT == "647300")||(GLACCOUNT == "647400"))
                    {
                            worksheet.Cells[cpt + 1, 11] = new Cell(Val);
                            worksheet.Cells[cpt + 1, 12] = new Cell(0);
                    }
             if ((GLACCOUNT == "453120"))
                    {
                            worksheet.Cells[cpt + 1, 11] = new Cell(0);
                            worksheet.Cells[cpt + 1, 12] = new Cell(Val);
                    }
            
        }

    }
}
