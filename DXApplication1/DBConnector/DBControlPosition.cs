using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.SqlClient;

namespace ALogic.DBConnector
{
    public static class DBControlPosition
    {
        public static object GetControlPosition(string nControl, int idUser)
        {
            SqlParameter nControlPar = new SqlParameter("nControl", nControl);
            SqlParameter idUserPar = new SqlParameter("idUser", idUser);

            return DBConnector.DBExecutor.SelectSchalar("select top 1 pos from ControlPosition where nControl = @nControl and idUser = @idUser", nControlPar, idUserPar);
        }

        public static void SaveControlPosition(string nControl, int idUser, int pos)
        {
            string sql = @"
if exists(select * from ControlPosition where nControl = @nControl and idUser = @idUser)
    update ControlPosition
        set pos = @pos
    where  nControl = @nControl and idUser = @idUser
else
    insert into ControlPosition(nControl, idUser, pos)
    values(@nControl, @idUser,  @pos)
";
            SqlParameter nControlPar = new SqlParameter("nControl", nControl);
            SqlParameter idUserPar = new SqlParameter("idUser", idUser);
            SqlParameter posPar = new SqlParameter("pos", pos);

            DBConnector.DBExecutor.ExecuteQuery(sql, nControlPar, idUserPar, posPar);
        }
    }
}
