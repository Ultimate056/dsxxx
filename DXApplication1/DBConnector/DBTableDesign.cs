using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;

namespace ALogic.DBConnector
{
    public class DBTableDesign
    {


        public static object getProperties(string tablename, int userid)
        {         
            SqlParameter nTable = new SqlParameter("@nTable", tablename);
            return DBExecutor.SelectSchalar(@"
if not exists(select * from sTableDesign where idUser = "+ userid  + @" and nTable = @nTable)
	select strDesign from sTableDesignDefault where nTable = @nTable;
else
	select strDesign from sTableDesign where idUser = " + userid+ @" and nTable = @nTable;
", nTable);          
        }

        public static bool saveProperties(string tablename, string strDesignpar, int userid)
        {
            //dbo.getKontrID() поменял на userid
            SqlParameter nTable = new SqlParameter("@nTable", tablename);
            SqlParameter strDesign = new SqlParameter("@strDesign", strDesignpar);
            return DBExecutor.ExecuteQuery(@"
if not exists(
select * from sTableDesign where idUser = "+userid+@" and nTable = @nTable)
insert into sTableDesign(nTable, strDesign, iduser) values (@nTable, @strDesign, "+userid+ @");
else
update sTableDesign set strDesign = @strDesign where idUser = "+userid+@" and nTable = @nTable;
", nTable, strDesign);

        }

        public static object getResetProperties(string tablename)
        {          
            SqlParameter nTable = new SqlParameter("@nTable", tablename);
           return DBExecutor.SelectSchalar(@"
select dbo.getTableDesignDefault(@nTable) as xmlstr
", nTable);           
        }

        public static void saveDefaultProperties(string tablename, string strDesignpar)
        {
            SqlParameter nTable = new SqlParameter("@nTable", tablename);
            SqlParameter strDesign = new SqlParameter("@strDesign", strDesignpar);

            DBExecutor.ExecuteQuery(@"
if not exists(select * from sTableDesignDefault where nTable = @nTable)
    insert into sTableDesignDefault(nTable, strDesign) values (@nTable, @strDesign);
else
    update sTableDesignDefault set strDesign = @strDesign where nTable = @nTable;
", nTable, strDesign);
        }

    }
}
