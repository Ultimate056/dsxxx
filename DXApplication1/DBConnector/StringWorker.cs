using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ALogic.DBConnector
{
    public static class StringWorker
    {
        public static string DelSymbols(string s, params char[] par)
        {
            foreach (var p in par)
                s = s.Replace(p.ToString(), "");

            return s;
        }

        public static string GetSubHard(string str, string mask, string separator)
        {
            string[] sn = mask.Split(new string[1] { separator }, StringSplitOptions.None);
            string sub1 = str.Split(new string[1] { sn[0] }, StringSplitOptions.None).Last();
            string sub2 = sub1.Split(new string[1] { sn[1] }, StringSplitOptions.None).First();

            return sub2.Trim();
        }
    }
}
