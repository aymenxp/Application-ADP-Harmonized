using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using HarmonizedApp.Manager.BLL;
using HarmonizedApp.Manager.BO;

namespace HarmonizedApp.Manager.BLL
{
    class clsGrossToNetBLL
    {
        // Extraction des données et split des colonnes

        //Méthode : Import 
        //Méthode : Export

        public void Import(int IdSoc,int Mois,int Annee,string FilePath)
        {
            HarmonizedApp.Manager.BLL.clsRubricStructureBLL rub = new clsRubricStructureBLL();
            List<clsRubricStructureBO> ListRubStruct = rub.GetListRubricsStructureByIdSoc(IdSoc);
            List<string> ListColumns = GetColumnListOfGrossToNetByIdSoc(IdSoc);
            List<string> ListValues = new List<string>();

            StreamReader sr = new StreamReader(FilePath);
            string line = null;

            // Retour => Field1,Field2,Field3
            // 
            foreach (clsRubricStructureBO r in ListRubStruct)
            {
                
            }


            while ((line = sr.ReadLine()) != null)
            {
                if (!(line.Equals("")))
                {
                //
                }
            }




        
        
        
        }

        public void Export()
        {

        }

        public List<string> GetColumnListOfGrossToNetByIdSoc(int IdSoc)
        {


            return null;

        }


    }
}
