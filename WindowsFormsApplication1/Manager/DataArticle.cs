using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace HarmonizedApp.Manager
{
    class DataArticle
    {
        private int mat;

        public int Mat
        {
            get { return mat; }
            set { mat = value; }
        }
        private string lib;

        public string Lib
        {
            get { return lib; }
            set { lib = value; }
        }
        private string prix;

        public string Prix
        {
            get { return prix; }
            set { prix = value; }
        }
    }
}
