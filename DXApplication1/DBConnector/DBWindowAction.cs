using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.SqlClient;
using System.Data;

namespace ALogic.DBConnector
{
    public static class DBWindowAction
    {
        public static DataRow GetWindowAction( int idWE)
        {
            SqlParameter idWEPar = new SqlParameter("idWE", idWE);
            string sql = "select * from sWE where idWE = @idWE";
            return DBExecutor.SelectRow(sql, idWEPar);

        }

        public static DataTable GetAllActions()
        {
            string sql = "select * from sWE";
            return DBExecutor.SelectTable(sql);
        }

        public static void Save(object nWE, object buttonText, object buttonImage, object number, object keyFlag, object skey)
        {
            string sql = "insert into sWE values(@nWE, @ButtonText, @ButtonImage, @Number, @keyFlag, @Skey)";

            SqlParameter nWEPar = new SqlParameter("nWE", nWE);
            SqlParameter buttonTextPar = new SqlParameter("buttonText", buttonText);
            SqlParameter buttonImagePar = new SqlParameter("buttonImage", buttonImage);
            SqlParameter numberPar = new SqlParameter("number", number);
            SqlParameter keyFlagPar = new SqlParameter("keyFlag", keyFlag);
            SqlParameter skeyPar = new SqlParameter("skey", skey);

            DBExecutor.ExecuteQuery(sql, nWEPar, buttonTextPar, buttonImagePar, numberPar, keyFlagPar, skeyPar);
        }

        public static void Update(object idWE, object nWE, object buttonText, object buttonImage, object number, object keyFlag, object skey)
        {
            string sql = @"
update sWE 
    set nWE = @nWE, 
        ButtonText = @ButtonText, 
        ButtonImage = @ButtonImage, 
        Number = @Number, 
        keyFlag = @keyFlag, 
        Skey = @Skey
where idWe = @idWe
";

            SqlParameter idWEPar = new SqlParameter("idWE", idWE);
            SqlParameter nWEPar = new SqlParameter("nWE", nWE);
            SqlParameter buttonTextPar = new SqlParameter("buttonText", buttonText);
            SqlParameter buttonImagePar = new SqlParameter("buttonImage", buttonImage);
            SqlParameter numberPar = new SqlParameter("number", number);
            SqlParameter keyFlagPar = new SqlParameter("keyFlag", keyFlag);
            SqlParameter skeyPar = new SqlParameter("skey", skey);

            DBExecutor.ExecuteQuery(sql,idWEPar, nWEPar, buttonTextPar, buttonImagePar, numberPar, keyFlagPar, skeyPar);
        }

        public static void Delete(object idWE)
        {
            string sql = "delete from sWE where idWE = @idWe";
            SqlParameter idWEPar = new SqlParameter("idWE", idWE);
            DBExecutor.ExecuteQuery(sql, idWEPar);
        }
    }
}
