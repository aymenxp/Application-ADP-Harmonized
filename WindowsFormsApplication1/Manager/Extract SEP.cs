using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using ExcelLibrary.SpreadSheet;

namespace HarmonizedApp.Manager
{
    class Extract_SEP
    {
        ServiceDAO service = new ServiceDAO();
        ContainSEP sep = new ContainSEP();


        public void ConstructSEP(string Date1, string Date2, string CodeSoc, string NomSoc, ProgressBar prog, int idSoc, int MoisPe,int annee)
        {
            string nomFichier = string.Empty;
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.DefaultExt = "xls";
            saveFileDialog1.FilterIndex = 2;
            saveFileDialog1.Filter = "fichiers excel (*.xls)|*.xls";
            saveFileDialog1.RestoreDirectory = true;
            saveFileDialog1.FileName = "P_SEP_" + CodeSoc + "_" + Date1 + "_" + Date2 + "_00_V2_0000_00000_FILE_NOE_" + NomSoc + "Paymentreport.xls";
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                try
                {

                    nomFichier = saveFileDialog1.FileName;
                    ExtraireDATA(nomFichier, CodeSoc, NomSoc, prog, Date1, Date2, idSoc, MoisPe,annee);
                }
                catch (Exception ex)
                {
                    ex.ToString();
                }
            }
        }

        public void ExtraireDATA(string nomFichier, string CodeSoc, string NomSoc, ProgressBar prog, string Date1, string Date2, int idSoc, int MoisPe,int annee)
        {
            prog.Visible = true;
            prog.Minimum = 0;
            //create new xls file
            Workbook workbook = new Workbook();
            Worksheet worksheet = new Worksheet("SEP");
            worksheet.Cells.ColumnWidth[0, 1] = 3000;
            worksheet = sep.EnteteSGNLigne1(worksheet);
            worksheet = sep.EnteteSVRLigne2(worksheet);
            worksheet = sep.EnteteSVRLigne3(worksheet);
            worksheet = sep.EnteteSVRLigne4(worksheet, NomSoc);
            worksheet = sep.EnteteSVRLigne5(worksheet, CodeSoc);
            worksheet = sep.EnteteSVRLigne6(worksheet, NomSoc);
            worksheet = sep.EnteteSVRLigne7(worksheet);
            worksheet = sep.EnteteSVRLigne8(worksheet, Date1, Date2);
            worksheet = sep.EnteteSVRLigne9(worksheet);
            worksheet = sep.EnteteSVRLigne10(worksheet);
            worksheet = sep.EnteteSVRLigne11(worksheet);
            worksheet = sep.EnteteSVRLigne12(worksheet, idSoc, MoisPe,annee);

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

    }
}
