using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace HarmonizedApp.Manager
{
    public class clsBD
    {
        public static string connection_string = @"Provider=Microsoft.Jet.OLEDB.4.0;" +
           @"Data source= " + Application.StartupPath + "\\ApplicationADP.mdb";


    }
}
