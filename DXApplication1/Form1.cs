﻿using ALogic.DBConnector;
using ALogic.Logic.SPR;
using DevExpress.Data.PivotGrid;
using DevExpress.XtraPivotGrid;
using DevExpress.XtraPrinting;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace DXApplication1
{
    public partial class Form1 : DevExpress.XtraEditors.XtraForm
    {
        delegate void UpdateStatusInfoDelegate(string msg, int part);

        public Form1()
        {
            InitializeComponent();
            DevExpress.XtraPivotGrid.PivotGridField.DefaultDecimalFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
            DevExpress.XtraPivotGrid.PivotGridField.DefaultDecimalFormat.FormatString = "N2";
            // This line of code is generated by Data Source Configuration Wizard
            //linqServerModeSource2.QueryableSource = new AForm.Forms.AReport.DataClassesSaleReportDataContext().tempdata_for_salereport.Where(s => s.iduser == User.Current.IdUser); 
            // This line of code is generated by Data Source Configuration Wizard

        }

        public void UpdateStatusInfo(string message, int part)
        {
            if (part == 0) { return; }
            var frac = toolStripProgressBar1.Maximum / part;
            var tvalue = toolStripProgressBar1.Value + frac;
            if (tvalue <= toolStripProgressBar1.Maximum)
            {
                toolStripProgressBar1.Value += frac;
            }

            toolStripStatusLabel1.Text = $"{message}";
        }

        public void GetReport(object StateInfo)
        {
            try
            {
                pivotGridControl1.DataSource = null;
                //linqServerModeSource2.Reload();

                Invoke(new Action(() => Cursor = Cursors.WaitCursor));
                Invoke(new Action(() => btnUpdate.Enabled = false));
                Invoke(new Action(() => btnSave.Enabled = false));

                var startTime = System.Diagnostics.Stopwatch.StartNew();

                Invoke(new Action(() => pivotGridControl1.BeginUpdate()));

                Invoke(new UpdateStatusInfoDelegate(UpdateStatusInfo), "Формирование отчета...", 3);

                DataClassesSaleReportDataContext dts = new DataClassesSaleReportDataContext();
                dts.CommandTimeout = 1800;

                string sqlstring = $@"exec up_Manage_SalesReport_U '{deStart.DateTime.ToString("yyyy.MM.dd")}', '{deEnd.DateTime.ToString("yyyy.MM.dd")}', {User.Current.IdUser}";
                DBExecutor.ExecuteQuery(sqlstring);

                //linqServerModeSource2.QueryableSource = new DataClassesSaleReportDataContext().tempdata_for_salereport.Where(s => s.iduser == User.Current.IdUser);
                //var nquery = dts.tempdata_for_salereport.Where(s => s.iduser == User.Current.IdUser);

                //DataTable dt = DBExecutor.SelectTable($"SELECT * FROM tempdata_for_salereport WHERE iduser = {User.Current.IdUser}");

                //linqServerModeSource2.QueryableSource = nquery;
                pivotGridControl1.DataSource = new DXApplication1.DataClassesSaleReportDataContext().tempdata_for_salereport.Where(x => x.iduser == User.Current.IdUser);
                Invoke(new UpdateStatusInfoDelegate(UpdateStatusInfo), $"Обработка...  строк", 3);

                //pivotGridControl1.DataSource = linqServerModeSource2;

                //pivotGridControl1.DataSource = DataSale.getSalePeriod(deStart.DateTime, deEnd.DateTime);
                //pivotGridControl1.DataSource = DataSale.getSalePeriodList(deStart.DateTime, deEnd.DateTime);

                Invoke(new Action(() => pivotGridControl1.EndUpdate()));

                startTime.Stop();
                var resultTime = startTime.Elapsed;

                // elapsedTime - строка, которая будет содержать значение затраченного времени
                string elapsedTime = String.Format("{0:00}:{1:00}:{2:00}.{3:000}", resultTime.Hours, resultTime.Minutes, resultTime.Seconds, resultTime.Milliseconds);

                Invoke(new UpdateStatusInfoDelegate(UpdateStatusInfo), $"Готово... Общее время выполнения {elapsedTime}", 3);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка формирования отчета: метод {GetCallerName()} {ex.Message}", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                Invoke(new Action(() => Cursor = Cursors.Default));
                Invoke(new Action(() => btnUpdate.Enabled = true));
                Invoke(new Action(() => btnSave.Enabled = true));
            }
        }

        /// <summary>
        /// Получение метода или свойства
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        static string GetCallerName([CallerMemberName]string name = "")
        {
            return name;
        }

        /// <summary>
        /// Вызов метода получения отчета в отдельном потоке
        /// </summary>
        public void UseThreadForReport()
        {
            toolStripProgressBar1.Value = 0;
            toolStripProgressBar1.Maximum = 100;

            try
            {
                ThreadPool.QueueUserWorkItem(GetReport);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка: метод {GetCallerName()} {ex.Message}");
            }
            finally
            {

            }
        }




        private void SalesReport_FormClosed(object sender, FormClosedEventArgs e)
        {
            linqServerModeSource2.QueryableSource = null;

            //    try
            //    {
            //        DBExecutor.ExecuteQuery($"delete from tempdata_for_salereport where iduser = {User.Current.IdUser}");
            //        linqServerModeSource2.QueryableSource = null;
            //    }
            //    catch (Exception ex)
            //    {
            //        MessageBox.Show("Ошибка удаления временных данных: " + ex.Message, "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //    }
            //    finally
            //    {
            //        this.Cursor = Cursors.Default;
            //    }
        }

        private void pivotGridControl1_FieldAreaChanged(object sender, DevExpress.XtraPivotGrid.PivotFieldEventArgs e)
        {
            Invoke(new UpdateStatusInfoDelegate(UpdateStatusInfo), "Готово...", 1);
        }

        private void SalesReport_Load(object sender, EventArgs e)
        {
            deStart.EditValue = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1);
            deEnd.EditValue = DateTime.Now.Date;
        }

        private void pivotGridControl1_CustomSummary(object sender, DevExpress.XtraPivotGrid.PivotGridCustomSummaryEventArgs e)
        {
            //string name = e.DataField.FieldName;
            //if (e.DataField.SummaryType == PivotSummaryType.Custom)
            //{
            //    IList list = e.CreateDrillDownDataSource();
            //    Hashtable ht = new Hashtable();
            //    for (int i = 0; i < list.Count; i++)
            //    {
            //        PivotDrillDownDataRow row = list[i] as PivotDrillDownDataRow;
            //        object v = row[name];
            //        if (v != null && v != DBNull.Value)
            //            ht[v] = null;
            //    }
            //    e.CustomValue = ht.Count;

            //}
        }

        private void pivotGridControl1_QueryException(object sender, PivotQueryExceptionEventArgs e)
        {
            MessageBox.Show(e.ErrorPanelText);
        }

        private void btnUpdate_Click(object sender, EventArgs e)
        {
            if(deStart.EditValue == null || deEnd.EditValue == null)
            {
                MessageBox.Show("Введите даты начала и конца срока продаж", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            UseThreadForReport();

            #region Закомментировано,пока не удалять !!!
            ////try
            ////{


            ////    toolStripProgressBar1.Value = 0;
            ////    toolStripProgressBar1.Maximum = 100;

            ////    pivotGridControl1.DataSource = null;
            ////    //linqServerModeSource2.Reload();

            ////    this.Cursor = Cursors.WaitCursor;
            ////    pivotGridControl1.BeginUpdate();

            ////    var startTime = System.Diagnostics.Stopwatch.StartNew();

            ////    Invoke(new UpdateStatusInfoDelegate(UpdateStatusInfo), "Начало формирования отчета...");

            ////    DataClassesSaleReportDataContext dts = new DataClassesSaleReportDataContext();
            ////    dts.CommandTimeout = 1800;


            ////    string sqlstring = $@"exec up_Manage_SalesReport_U '{deStart.DateTime.ToString("yyyy.MM.dd")}', '{deEnd.DateTime.ToString("yyyy.MM.dd")}', {User.Current.IdUser}";
            ////    DBExecutor.ExecuteQuery(sqlstring);

            ////    Invoke(new UpdateStatusInfoDelegate(UpdateStatusInfo), "Данные получены...");

            ////    //linqServerModeSource2.QueryableSource = new DataClassesSaleReportDataContext().tempdata_for_salereport.Where(s => s.iduser == User.Current.IdUser);
            ////    linqServerModeSource2.QueryableSource = dts.tempdata_for_salereport.Where(s => s.iduser == User.Current.IdUser);
            ////    pivotGridControl1.DataSource = linqServerModeSource2;

            ////    //upd MuhinAN 21/06/2022 закомменчено до выяснения тормозов
            ////    //linqServerModeSource2.Reload();

            ////    //DataSale.getSalePeriodList(deStart.DateTime, deEnd.DateTime, User.Current.IdUser);

            ////    //string sqlstring = @"exec up_Manage_SalesReport_U '" + deStart.DateTime.ToString("yyyy.MM.dd") + "', '" + deEnd.DateTime.ToString("yyyy.MM.dd") + "', " + User.Current.IdUser.ToString();

            ////    //linqServerModeSource2.ElementType = typeof(tempdata_for_salereport);
            ////    //DataClassesSaleReportDataContext dts = new DataClassesSaleReportDataContext();
            ////    //dts.CommandTimeout = 1800;
            ////    //linqServerModeSource2.QueryableSource = dts.ExecuteQuery(typeof(tempdata_for_salereport), sqlstring).AsQueryable();
            ////    //linqServerModeSource2.KeyExpression = "tm_name";

            ////    //linqServerModeSource2.KeyExpression = "iduser";

            ////    //////// This line of code is generated by Data Source Configuration Wizard
            ////    //linqServerModeSource2.QueryableSource = new DataClassesSaleReportDataContext().tempdata_for_salereport.Where(s => s.iduser == User.Current.IdUser);
            ////    //DataClassesManageReportDataContext dctnx = new DataClassesManageReportDataContext();
            ////    //dctnx.up_Manage_SalesReport(deStart.DateTime, deEnd.DateTime, User.Current.IdUser);
            ////    //pivotGridControl1.DataSource = linqServerModeSource2;
            ////    //end upd MuhinAN 21/06/2022

            ////    //pivotGridControl1.DataSource = DataSale.getSalePeriod(deStart.DateTime, deEnd.DateTime);
            ////    //pivotGridControl1.DataSource = DataSale.getSalePeriodList(deStart.DateTime, deEnd.DateTime);

            ////    startTime.Stop();
            ////    var resultTime = startTime.Elapsed;

            ////    // elapsedTime - строка, которая будет содержать значение затраченного времени
            ////    string elapsedTime = String.Format("{0:00}:{1:00}:{2:00}.{3:000}",
            ////        resultTime.Hours,
            ////        resultTime.Minutes,
            ////        resultTime.Seconds,
            ////        resultTime.Milliseconds);

            ////    pivotGridControl1.EndUpdate();

            ////    Invoke(new UpdateStatusInfoDelegate(UpdateStatusInfo), "Готово...");
            ////}
            ////catch (Exception ex)
            ////{
            ////    MessageBox.Show("Ошибка формирования отчета: " + ex.Message, "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Error);
            ////}
            ////finally
            ////{
            ////    this.Cursor = Cursors.Default;
            ////}
            #endregion!!!
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                var opt = new XlsxExportOptionsEx();

                opt.AllowGrouping = DevExpress.Utils.DefaultBoolean.False;
                opt.AllowFixedColumnHeaderPanel = DevExpress.Utils.DefaultBoolean.False;
                opt.TextExportMode = DevExpress.XtraPrinting.TextExportMode.Value;
                opt.RawDataMode = true;

                pivotGridControl1.ExportToXlsx(saveFileDialog1.FileName + ".xlsx", opt);
            }
        }

        private void btnSettings_Click(object sender, EventArgs e)
        {
            //fmSettings set = new fmSettings();
            //set.gridobject = pivotGridControl1;
            //set.appname = this.Name;
            //set.StartPosition = FormStartPosition.CenterParent;
            //set.ShowDialog();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            deStart.EditValue = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1);
            deEnd.EditValue = DateTime.Now.Date;
        }

        //private void button1_Click(object sender, EventArgs e)
        //{
        //    var tbl = DataSale.getSalePeriod(deStart.DateTime, deEnd.DateTime);

        //    var ds = @"C:\Users\MUHINAN.ARKONA\Source\Repos\VS_ARKONA\AForm\test.xls";

        //    if (File.Exists(ds)) { File.Delete(ds); }

        //    using (OleDbConnection conn = new OleDbConnection($@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={ds};Extended Properties='Excel 8.0;HDR=Yes'"))
        //    {
        //        conn.Open();

        //        OleDbCommand cmd = new OleDbCommand(@"CREATE TABLE [Sheet1]([Column1] string, [Column2] numeric)", conn);
        //        cmd.ExecuteNonQuery();

        //        foreach (DataRow r in tbl.Rows)
        //        {
        //            var br = r["tm_name"];
        //            var sumsale = r["sumsale"];
        //            cmd.CommandText = $@"INSERT INTO [Sheet1](Column1, Column2) values(?, ?)";
        //            cmd.Parameters.AddWithValue("@tmname", OleDbType.VarChar).Value = br;
        //            cmd.Parameters.AddWithValue("@sumsale", OleDbType.Decimal).Value = sumsale;
        //            cmd.ExecuteNonQuery();
        //        }

        //        conn.Close();
        //    }

        //}
    }
}
