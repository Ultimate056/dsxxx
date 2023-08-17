using ALogic.Logic.Reload1C;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;

namespace ALogic.DBConnector
{
    public static class DBGroupSaver
    {
        const int CountRowsPerOne = 100000;
        
        private static StringBuilder _query;
        private static int _countRows;

        static DBGroupSaver()
        {
            _query = new StringBuilder(100000);
            _countRows = 0;
        }

        public static void AddRow( string query )
        {
            _query.Append(query);
            _countRows++;

            if (_countRows > CountRowsPerOne)
            {
                SaveAll();
            }
        }

        public static void SaveAll()
        {
            try
            {
                DBExecutor.ExecuteTranzactionQuery(_query.ToString());
                _query.Clear();
                _countRows = 0;
            }
            catch (Exception ex)
            {
                string ee = ex.Message;
                string x = _query.ToString();
                _query.Clear();
                _countRows = 0;
                UniLogger.WriteLog("", 1, "Ошибка метода SaveAll: " + ex.Message);
                return;
            }
            finally
            {
                GC.Collect();
            }
        }

        public static void AddToDataTable()
        {
            //insert into PriceOnlineTemp (idUser, idKontr,DatePrice,idCur,idTm,nTov,ntovsupplier,kol,priceCur,idTovOEM,brend, nBrandSupplier) 
            //values('552632', 556643, '9.13.2021', 0, 0, '', 'Паста полировальная высокоабразивная, №1 Fast Cut Compound, 1 кг', '2', '2362.50', '09374', '3M', '3M')

            DataTable dt = new DataTable();
            dt.Columns.AddRange(new DataColumn[12]
                        { new DataColumn("idUser", typeof(int)), new DataColumn("idKontr", typeof(int)), new DataColumn("DatePrice",typeof(string)),
                                new DataColumn("idCur", typeof(int)), new DataColumn("idTm", typeof(int)), new DataColumn("nTov", typeof(string)),
                                new DataColumn("ntovsupplier", typeof(string)), new DataColumn("kol",typeof(int)), new DataColumn("priceCur",typeof(string)),
                                new DataColumn("idTovOEM",typeof(string)), new DataColumn("brend",typeof(string)), new DataColumn("nBrandSupplier",typeof(string))
                        });


        }
    }
}
