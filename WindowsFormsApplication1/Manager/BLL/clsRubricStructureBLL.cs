using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using HarmonizedApp.Manager.BO;
using HarmonizedApp.Manager.DAL;
using HarmonizedApp.Manager;

namespace HarmonizedApp.Manager.BLL
{
    public class clsRubricStructureBLL
    {
        /// <summary>
        /// Retourne les ordres des colonnes select Col1,Col2,.....,Coln
        /// </summary>
        /// <returns></returns>
        public string PrepareSqlQuery(int IdSoc)
        {
            // Get Rubrique ordred by order parameter
            // 
            clsRubricStructureDAL Dal = new clsRubricStructureDAL();
            string SqlQuery="";
            int index = 0;
            List<clsRubricStructureBO> ListRubrics = Dal.GetListRubricsByIdSoc(IdSoc);
         
            foreach (clsRubricStructureBO r in ListRubrics)
            {
                SqlQuery = SqlQuery + r.NomColonne ;

                if (index != ListRubrics.Count() - 1)
                    SqlQuery = SqlQuery + ",";
            }


            string GrossToNetTableName = "";

            GrossToNetTableName = ManagerDAO.GetInstance().GetDetailsSociete(IdSoc).GrossToNetTableName;

           return " Select " + SqlQuery + " From " + GrossToNetTableName;

        }

        public List<clsRubricStructureBO> GetListRubricsStructureByIdSoc(int IdSoc)
        {
            // Get Rubrique ordred by order parameter

            clsRubricStructureDAL Dal = new clsRubricStructureDAL();
           return  Dal.GetListRubricsByIdSoc(IdSoc);
        }

    }
}
