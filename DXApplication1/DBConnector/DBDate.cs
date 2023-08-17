using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ALogic.DBConnector
{
    public static class DBDate
    {
        public static DateTime Date { get; set; }
        public static string StrDate { get { return Date.Month.ToString() + "." + Date.Day.ToString() + "." + Date.Year.ToString(); } }
    }
}
