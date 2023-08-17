using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.SqlClient;
using System.Data;

namespace ALogic.DBConnector
{
    public static class DBAppParam
    {
        public static object GetAppParamValue(int idAppParam)
        {
            SqlParameter idAppParampar = new SqlParameter("idAppParam", idAppParam);
            return DBExecutor.SelectSchalar("select value_param from application_param where id_param = @idAppParam", idAppParampar);
        }

        public static DataRow SetAppParamValue(int idAppParam, object value)
        {
            SqlParameter idAppParampar = new SqlParameter("idAppParam", idAppParam);
            SqlParameter valuepar = new SqlParameter("value", value);
            return DBExecutor.SelectRow("update application_param set value_param = @value where id_param = @idAppParam", idAppParampar, valuepar);
        }

        public static object GetAppParamYN(int idAppParam)
        {
            SqlParameter idAppParampar = new SqlParameter("idAppParam", idAppParam);
            return DBExecutor.SelectSchalar("select yes_no from application_param where id_param = @idAppParam", idAppParampar);
        }

        public static DataRow SetAppParamYN(int idAppParam, object value)
        {
            SqlParameter idAppParampar = new SqlParameter("idAppParam", idAppParam);
            SqlParameter valuepar = new SqlParameter("value", value);
            return DBExecutor.SelectRow("update application_param set yes_no = @value where id_param = @idAppParam", idAppParampar, valuepar);
        }

    }
}
