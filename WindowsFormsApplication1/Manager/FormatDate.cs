using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace HarmonizedApp.Manager
{
    class FormatDate
    {
        public string LA_DATE(string laDate)
        {
            string DATE_1 = string.Empty;
            string YEAR, MONTH, DAY = string.Empty;
            DATE_1 = laDate.Substring(0, 10);
            YEAR = DATE_1.Substring(6, 4);
            MONTH = DATE_1.Substring(3, 2);
            DAY = DATE_1.Substring(0, 2);
            DATE_1 = YEAR + MONTH + DAY;
            return DATE_1;
        }

        public string LA_DATE_Reporting(string laDate)
        {
            string DATE_1 = string.Empty;
            string YEAR, MONTH, DAY = string.Empty;
            DATE_1 = laDate.Substring(0, 10);
            YEAR = DATE_1.Substring(6, 4);
            MONTH = DATE_1.Substring(3, 2);
            DAY = DATE_1.Substring(0, 2);
            DATE_1 = YEAR + "-" + MONTH + "-" + DAY;
            return DATE_1;
        }

        public string LA_DATE_ADP(string laDate)
        {
            string DATE_1 = string.Empty;
            string YEAR, MONTH, DAY = string.Empty;
            DATE_1 = laDate.Substring(0, 8);
            YEAR = DATE_1.Substring(0, 4);
            MONTH = DATE_1.Substring(4, 2);
            DAY = DATE_1.Substring(6, 2);
            DATE_1 = YEAR + "/" + MONTH + "/" + DAY;
            return DATE_1;
        }

        public string ExtraireMoisPe(string laDate)
        {
            string DATE_1 = string.Empty;
            string MONTH = string.Empty;
            string Mois=string.Empty;
            DATE_1 = laDate.Substring(0, 10);
            MONTH = DATE_1.Substring(3, 2);
            if (MONTH.Length == 1)
            {
                Mois = "0" + MONTH;
            }
            else
            {
                Mois = MONTH;
            }
            return Mois;
        }

        public string DateAmeric(string DateEmb)
        {
            string DateF = string.Empty;
            if ((DateEmb == "")||(DateEmb=="Vide"))
            {
                DateF = "";
            }
            else
            {              
                string Year = DateEmb.Substring(6, 2);
                string Month = DateEmb.Substring(3, 2);
                string Day = DateEmb.Substring(0, 2);
               
                    if (Year != "")
                    {
                        int annee = int.Parse(Year);
                        if ((annee <= 99) && (annee >= 50))
                        {
                            DateF = "19" + Year + "/";
                        }
                        else
                        {
                            DateF = "20" + Year + "/";
                        }
                        DateF = DateF + Month + "/" + Day;
                    }              
            }
            return DateF;
        }


        public int GetMonthByName(string MonthName)
        {
            switch (MonthName.ToLower())
            {
                case "janvier":
                    return 1;
                case "février":
                    return 2;
                case "mars":
                    return 3;
                case "avril":
                    return 4;
                case "mai":
                    return 5;
                case "juin":
                    return 6;
                case "juillet":
                    return 7;
                case "août":
                    return 8;
                case "septembre":
                    return 9;
                case "octobre":
                    return 10;
                case "novembre":
                    return 11;
                case "décembre":
                    return 12;
                default:
                    return 0;
            }
        }

    }
}
