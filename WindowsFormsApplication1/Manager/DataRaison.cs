using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace HarmonizedApp.Manager
{
    class DataRaison
    {
        private int num;

        public int Num
        {
            get { return num; }
            set { num = value; }
        }
        private string societe;

        public string Societe
        {
            get { return societe; }
            set { societe = value; }
        }

        private string codeSoc;

        public string CodeSoc
        {
            get { return codeSoc; }
            set { codeSoc = value; }
        }

        private string _GrossToNetTableName;

        public string GrossToNetTableName
        {
            get { return _GrossToNetTableName; }
            set { _GrossToNetTableName = value; }
        }



    }
}
