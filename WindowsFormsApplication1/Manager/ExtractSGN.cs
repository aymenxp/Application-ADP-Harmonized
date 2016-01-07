using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using ExcelLibrary.SpreadSheet;

namespace HarmonizedApp.Manager
{
    class ExtractSGN
    {

        ServiceDAO service = new ServiceDAO();
        EnteteSGN sgn = new EnteteSGN();


        public void ConstructSGN(string Date1, string Date2, string CodeSoc, string NomSoc, ProgressBar prog, int idSoc, int MoisPe,int AnnePe)
        {
            string nomFichier = string.Empty;
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.DefaultExt = "xls";
            saveFileDialog1.FilterIndex = 2;
            saveFileDialog1.Filter = "fichiers excel (*.xls)|*.xls";
            saveFileDialog1.RestoreDirectory = true;
            saveFileDialog1.FileName = "P_SGN_" + CodeSoc + "_" + Date1 + "_" + Date2 + "_00_V2_0000_00000_FILE_NOE_" + NomSoc + "Grosstonetreport.xls";
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                try
                {

                    nomFichier = saveFileDialog1.FileName;
                    ExtraireDATA(nomFichier, CodeSoc, NomSoc, prog, Date1, Date2, idSoc, MoisPe,AnnePe);
                }
                catch (Exception ex)
                {
                    ex.ToString();
                }
            }
        }

        public void ExtraireDATA(string nomFichier, string CodeSoc, string NomSoc, ProgressBar prog, string Date1, string Date2, int idSoc, int MoisPe,int AnnePe)
        {
            prog.Visible = true;
            prog.Minimum = 0;
            //create new xls file
            Workbook workbook = new Workbook();
            Worksheet worksheet = new Worksheet("SGN");
            worksheet.Cells.ColumnWidth[0, 1] = 3000;
            worksheet = sgn.EnteteSGNLigne1(worksheet);
            worksheet = sgn.EnteteSVRLigne2(worksheet);
            worksheet = sgn.EnteteSVRLigne3(worksheet);
            worksheet = sgn.EnteteSVRLigne4(worksheet,NomSoc);
            worksheet = sgn.EnteteSVRLigne5(worksheet,CodeSoc);
            worksheet = sgn.EnteteSVRLigne6(worksheet, NomSoc);
            worksheet = sgn.EnteteSVRLigne7(worksheet);
            worksheet = sgn.EnteteSVRLigne8(worksheet, Date1, Date2);
            worksheet = sgn.EnteteSVRLigne9(worksheet);
            worksheet = sgn.EnteteSVRLigne10(worksheet);
            EcritureC(idSoc, worksheet, MoisPe,AnnePe);
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

        public void EcritureC(int idSoc, Worksheet worksheet, int MoisPe,int AnnePe)
        {
            if (idSoc == 2)
            {
                worksheet = sgn.EnteteSVRLigne11(worksheet);
                worksheet = sgn.EnteteSVRLigne12(worksheet, idSoc, MoisPe,AnnePe);
            }
            if (idSoc == 5)
            {
                worksheet = sgn.EnteteSVRLigne11A002(worksheet);
                worksheet = sgn.EnteteSVRLigne12A002(worksheet, idSoc, MoisPe,AnnePe);
            }
            if (idSoc == 8)
            {
                worksheet = sgn.EnteteSVRLigne11A003(worksheet);
                worksheet = sgn.EnteteSVRLigne12A003(worksheet, idSoc, MoisPe, AnnePe);
            }
            if (idSoc == 7)
            {
                worksheet = sgn.EnteteSVRLigne11A004(worksheet);
                worksheet = sgn.EnteteSVRLigne12A004(worksheet, idSoc, MoisPe, AnnePe);
            }
            // Sony Ericson
            if (idSoc == 14)
            {
                worksheet = sgn.EnteteSVRLigne11A021(worksheet);
                worksheet = sgn.EnteteSVRLigne12A021(worksheet, idSoc, MoisPe, AnnePe);
            }

            if (idSoc == 9)
            {
                worksheet = sgn.EnteteSVRLigne11A010(worksheet);
                worksheet = sgn.EnteteSVRLigne12A010(worksheet, idSoc, MoisPe, AnnePe);
            }
            if (idSoc == 10)
            {
                worksheet = sgn.EnteteSVRLigne11A011(worksheet);
                worksheet = sgn.EnteteSVRLigne12A011(worksheet, idSoc, MoisPe, AnnePe);
            }
            if (idSoc == 3)
            {
                worksheet = sgn.EnteteSVRLigne11A014(worksheet);
                worksheet = sgn.EnteteSVRLigne12A014(worksheet, idSoc, MoisPe, AnnePe);
            }
            if (idSoc == 12)
            {
                worksheet = sgn.EnteteSVRLigne11A015(worksheet);
                worksheet = sgn.EnteteSVRLigne12A015(worksheet, idSoc, MoisPe, AnnePe);
            }
            if (idSoc == 11)
            {
                worksheet = sgn.EnteteSVRLigne11A013(worksheet);
                worksheet = sgn.EnteteSVRLigne12A013(worksheet, idSoc, MoisPe, AnnePe);
            }
            if (idSoc == 4)
            {
                worksheet = sgn.EnteteSVRLigne11A020(worksheet);
                worksheet = sgn.EnteteSVRLigne12A020(worksheet, idSoc, MoisPe, AnnePe);
            }
            if (idSoc == 13)
            {
                worksheet = sgn.EnteteSVRLigne11A016(worksheet);
                worksheet = sgn.EnteteSVRLigne12A016(worksheet, idSoc, MoisPe, AnnePe);
            }
         
        }

    }
}
