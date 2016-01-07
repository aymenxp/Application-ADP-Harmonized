using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using ExcelLibrary.SpreadSheet;

namespace HarmonizedApp.Manager
{
    class TraitementGRFile
    {

        ServiceDAO service = new ServiceDAO();
        EnteteSGR entete = new EnteteSGR();


        public void ExtraireDataSGL(string TypeFich, string NomFich, string Date1, string Date2, string DateJour, string CodeSoc, string NomSoc, int idSoc, int MoisPe, int anneePe, string TypeTri)
        {
            string nomFichier = string.Empty;
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.DefaultExt = "xls";
            saveFileDialog1.FilterIndex = 2;
            saveFileDialog1.Filter = "fichiers excel (*.xls)|*.xls";
            saveFileDialog1.RestoreDirectory = true;
            saveFileDialog1.FileName = "" + TypeFich + "" + CodeSoc + "_" + Date1 + "_" + Date2 + "_00_V2_0000_00000_FILE_NOE_" + NomSoc + "" + NomFich + "";
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                try
                {

                    nomFichier = saveFileDialog1.FileName;
                    ExtraireDATA(nomFichier, CodeSoc, NomSoc, Date1, Date2, DateJour, idSoc, MoisPe, anneePe, TypeTri);
                }
                catch (Exception ex)
                {
                    ex.ToString();
                }
            }
        }

        public void ExtraireDataOPR8(string TypeFich,string NomFich, string Date1, string Date2,string CodeSoc, string NomSoc, int idSoc, int MoisPe1, int anneePe1,int MoisPe2,int AnneePe2)
        {
            string nomFichier = string.Empty;
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.DefaultExt = "xls";
            saveFileDialog1.FilterIndex = 2;
            saveFileDialog1.Filter = "fichiers excel (*.xls)|*.xls";
            saveFileDialog1.RestoreDirectory = true;
            saveFileDialog1.FileName = "" + TypeFich + "" + CodeSoc + "_" + Date1 + "_" + Date2 + "_" + NomSoc + "" + NomFich + "";
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    nomFichier = saveFileDialog1.FileName;
                    ExtraireOPR8(nomFichier, CodeSoc, NomSoc, Date1, Date2,  idSoc, MoisPe1, anneePe1,MoisPe2,AnneePe2);
                }
                catch (Exception ex)
                {
                    ex.ToString();
                }
            }
        }

        public void ExtraireDataOPR11(string TypeFich, string NomFich, string Date1,string CodeSoc, string NomSoc, int idSoc, int MoisPe1, int anneePe1)
        {
            string nomFichier = string.Empty;
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.DefaultExt = "xls";
            saveFileDialog1.FilterIndex = 2;
            saveFileDialog1.Filter = "fichiers excel (*.xls)|*.xls";
            saveFileDialog1.RestoreDirectory = true;
            saveFileDialog1.FileName = "" + TypeFich + "" + CodeSoc + "_" + Date1 + "_" + NomSoc + "" + NomFich + "";
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    nomFichier = saveFileDialog1.FileName;
                    ExtraireOPR11(nomFichier, CodeSoc, NomSoc, Date1, idSoc, MoisPe1, anneePe1);
                }
                catch (Exception ex)
                {
                    ex.ToString();
                }
            }
        }

        public void ExtraireDataOPR17(string TypeFich, string NomFich, string Date1, string CodeSoc, string NomSoc, int idSoc, int MoisPe1, int anneePe1)
        {
            string nomFichier = string.Empty;
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.DefaultExt = "xls";
            saveFileDialog1.FilterIndex = 2;
            saveFileDialog1.Filter = "fichiers excel (*.xls)|*.xls";
            saveFileDialog1.RestoreDirectory = true;
            saveFileDialog1.FileName = "" + TypeFich + "" + CodeSoc + "_" + Date1 + "_" + NomSoc + "" + NomFich + "";
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    nomFichier = saveFileDialog1.FileName;
                    ExtraireOPR17(nomFichier, CodeSoc, NomSoc, Date1, idSoc, MoisPe1, anneePe1);
                }
                catch (Exception ex)
                {
                    ex.ToString();
                }
            }
        }

        public void ExtraireDataOPR15(string TypeFich, string NomFich, string Date1, string CodeSoc, string NomSoc, int idSoc, int MoisPe1, int anneePe1)
        {
            string nomFichier = string.Empty;
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.DefaultExt = "xls";
            saveFileDialog1.FilterIndex = 2;
            saveFileDialog1.Filter = "fichiers excel (*.xls)|*.xls";
            saveFileDialog1.RestoreDirectory = true;
            saveFileDialog1.FileName = "" + TypeFich + "" + CodeSoc + "_" + Date1 + "_" + NomSoc + "" + NomFich + "";
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    nomFichier = saveFileDialog1.FileName;
                    ExtraireOPR15(nomFichier, CodeSoc, NomSoc, Date1, idSoc, MoisPe1, anneePe1);
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
            Worksheet worksheet = new Worksheet("SGR");
            worksheet.Cells.ColumnWidth[0, 1] = 3000;
            worksheet = entete.EnteteSVRLigne1(worksheet);
            worksheet = entete.EnteteSVRLigne2(worksheet, Date1, Date2, DateJour, NomSoc, CodeSoc, TypeTri, idSoc, MoisPe, AnneePe);
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

        public void ExtraireOPR8(string nomFichier, string CodeSoc, string NomSoc, string Date1, string Date2, int idSoc, int MoisPe1, int AnneePe1, int MoisPe2,int AnneePe2)
        {
            //create new xls file
            Workbook workbook = new Workbook();
            Worksheet worksheet = new Worksheet("OPR8");
            worksheet.Cells.ColumnWidth[0, 1] = 3000;
            worksheet = entete.EnteteOPR8Ligne1(worksheet,NomSoc);
            worksheet = entete.EnteteOPR8Ligne2(worksheet);
            worksheet = entete.EnteteOPR8Ligne3(worksheet,MoisPe1);
            worksheet = entete.EnteteOPR8Ligne4(worksheet, AnneePe1);
            worksheet = entete.EnteteOPR8Ligne5(worksheet);
            worksheet = entete.EnteteOPR8Ligne6(worksheet,MoisPe1,AnneePe1,MoisPe2,AnneePe2);
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

        public void ExtraireOPR11(string nomFichier, string CodeSoc, string NomSoc, string Date1,int idSoc, int MoisPe1, int AnneePe1)
        {
            //create new xls file
            Workbook workbook = new Workbook();
            Worksheet worksheet = new Worksheet("OPR11");
            worksheet.Cells.ColumnWidth[0, 1] = 3000;
            worksheet = entete.EnteteOPR8Ligne1(worksheet, NomSoc);
            worksheet = entete.EnteteOPR11Ligne2(worksheet);
            worksheet = entete.EnteteOPR8Ligne3(worksheet, MoisPe1);
            worksheet = entete.EnteteOPR8Ligne4(worksheet, AnneePe1);
            worksheet = entete.EnteteOPR11Ligne5(worksheet);
            worksheet = entete.EnteteOPR11Ligne6(worksheet, MoisPe1, AnneePe1);
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

        public void ExtraireOPR17(string nomFichier, string CodeSoc, string NomSoc, string Date1, int idSoc, int MoisPe1, int AnneePe1)
        {
            //create new xls file
            Workbook workbook = new Workbook();
            Worksheet worksheet = new Worksheet("OPR17");
            worksheet.Cells.ColumnWidth[0, 1] = 3000;
            worksheet = entete.EnteteOPR8Ligne1(worksheet, NomSoc);
            worksheet = entete.EnteteOPR17Ligne2(worksheet);
            worksheet = entete.EnteteOPR8Ligne3(worksheet, MoisPe1);
            worksheet = entete.EnteteOPR8Ligne4(worksheet, AnneePe1);
            worksheet = entete.EnteteOPR17Ligne5(worksheet);
            worksheet = entete.EnteteOPR17Ligne6(worksheet, MoisPe1, AnneePe1);
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

        public void ExtraireOPR15(string nomFichier, string CodeSoc, string NomSoc, string Date1, int idSoc, int MoisPe1, int AnneePe1)
        {
            //create new xls file
            Workbook workbook = new Workbook();
            Worksheet worksheet = new Worksheet("OPR15");
            worksheet.Cells.ColumnWidth[0, 1] = 3000;
            worksheet = entete.EnteteOPR8Ligne1(worksheet, NomSoc);
            worksheet = entete.EnteteOPR15Ligne2(worksheet);
            worksheet = entete.EnteteOPR8Ligne3(worksheet, MoisPe1);
            worksheet = entete.EnteteOPR8Ligne4(worksheet, AnneePe1);
            worksheet = entete.EnteteOPR15Ligne5(worksheet);
            worksheet = entete.EnteteOPR15Ligne6(worksheet, MoisPe1, AnneePe1);
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
