using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using ExcelLibrary.SpreadSheet;
using System.Data;

namespace HarmonizedApp.Manager
{
    class TraitementSVR
    {
        ServiceDAO service = new ServiceDAO();
        Entete_SVR entete = new Entete_SVR();


        public void ExtraireDataSVR(string Date1, string Date2, string CodeSoc,string NomSoc, ProgressBar prog,int idSoc,int MoisPe,int anneePe)
        {
            string nomFichier = string.Empty;
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.DefaultExt = "xls";
            saveFileDialog1.FilterIndex = 2;
            saveFileDialog1.Filter = "fichiers excel (*.xls)|*.xls";
            saveFileDialog1.RestoreDirectory = true;
            saveFileDialog1.FileName = "P_SVR_" + CodeSoc + "_" + Date1 + "_" + Date2 + "_00_V2_0000_00000_FILE_NOE_" + NomSoc + "Summaryvariancereport.xls";
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                try
                {

                    nomFichier = saveFileDialog1.FileName;
                    ExtraireDATA(nomFichier, CodeSoc, NomSoc,prog,Date1,Date2,idSoc,MoisPe,anneePe);
                }
                catch (Exception ex)
                {
                    ex.ToString();
                }
            }
        }

        public void ExtraireDATA(string nomFichier, string CodeSoc, string NomSoc,ProgressBar prog,string Date1,string Date2,int idSoc,int MoisPe,int AnneePe)
        {
            prog.Visible = true;
            prog.Minimum = 0;
            //create new xls file
            Workbook workbook = new Workbook();
            Worksheet worksheet = new Worksheet("SVR");
            worksheet.Cells.ColumnWidth[0, 1] = 3000;
            worksheet = entete.EnteteSVRLigne1(worksheet);
            worksheet = entete.EnteteSVRLigne2(worksheet);
            worksheet = entete.EnteteSVRLigne3(worksheet);
            worksheet = entete.EnteteSVRLigne4(worksheet,NomSoc);
            worksheet = entete.EnteteSVRLigne5(worksheet, CodeSoc);
            worksheet = entete.EnteteSVRLigne6(worksheet,NomSoc);
            worksheet = entete.EnteteSVRLigne7(worksheet);
            worksheet = entete.EnteteSVRLigne8(worksheet,Date1,Date2);
            worksheet = entete.EnteteSVRLigne9(worksheet);
            worksheet = entete.EnteteSVRLigne10(worksheet);
            worksheet = entete.EnteteSVRLigne11(worksheet);
            worksheet = entete.EnteteSVRLigne12(worksheet);
            worksheet = entete.EnteteSVRLigne13(worksheet, idSoc, MoisPe);           
            EcritureC(idSoc, worksheet, MoisPe, AnneePe);    
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

        public void EcritureC(int idSoc, Worksheet worksheet, int MoisPe, int AnneePe)
        {
    //        if (idSoc == 2)
    //        {
                worksheet = entete.EnteteSVRLigne14(worksheet, idSoc, MoisPe, AnneePe);
    //        }
        }


    }
}
