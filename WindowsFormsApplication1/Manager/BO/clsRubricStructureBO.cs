using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace HarmonizedApp.Manager.BO
{
    public class clsRubricStructureBO
    {
        public int Id { get; set; }
        public int IdSoc { get; set; }
        public string Label { get; set; }
        public  string NomColonne { get; set; }
        public string ColonneType { get; set; }
        public int StartString { get; set; }
        public int EndString { get; set; }
        public string ConvertOut { get; set; }
        public string ReplaceString { get; set; }
        public int Order { get; set; }

    }
}
