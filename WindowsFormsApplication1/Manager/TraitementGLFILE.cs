using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using ExcelLibrary.SpreadSheet;

namespace HarmonizedApp.Manager
{
    class TraitementGLFILE
    {

        ServiceDAO service = new ServiceDAO();
        EnteteSGL entete = new EnteteSGL();




        public void ExtraireDataSGL(string TypeFich, string NomFich,string Date1,string Date2,string DateJour, string CodeSoc, string NomSoc,int idSoc, int MoisPe, int anneePe,string TypeTri)
        {
            string nomFichier = string.Empty;
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.DefaultExt = "xls";
            saveFileDialog1.FilterIndex = 2;
            saveFileDialog1.Filter = "fichiers excel (*.xls)|*.xls";
            saveFileDialog1.RestoreDirectory = true;
            saveFileDialog1.FileName = ""+TypeFich+"" + CodeSoc + "_" + Date1 + "_" + Date2 + "_00_V2_0000_00000_FILE_NOE_" + NomSoc + ""+NomFich+"";
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                try
                {

                    nomFichier = saveFileDialog1.FileName;
                    ExtraireDATA(nomFichier, CodeSoc, NomSoc,Date1, Date2,DateJour, idSoc, MoisPe, anneePe,TypeTri);
                }
                catch (Exception ex)
                {
                    ex.ToString();
                }
            }
        }

        public void ExtraireDATA(string nomFichier, string CodeSoc, string NomSoc, string Date1, string Date2, string DateJour, int idSoc, int MoisPe, int AnneePe, string TypeTri)
        {
            //create new xls file
            Workbook workbook = new Workbook();
            Worksheet worksheet = new Worksheet("SGL");
            worksheet.Cells.ColumnWidth[0, 1] = 3000;
            worksheet = entete.EnteteSVRLigne1(worksheet);
            worksheet = entete.EnteteSVRLigne2(worksheet, Date1,Date2,DateJour,NomSoc, CodeSoc, TypeTri,idSoc,MoisPe,AnneePe);
            workbook.Worksheets.Add(worksheet);
            workbook.Save(nomFichier);
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
