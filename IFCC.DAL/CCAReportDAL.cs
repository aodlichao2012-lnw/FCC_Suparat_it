using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.Data;


namespace IFCC.DAL
{
    public class CCAReportDAL : BaseDAL
    {

        #region + Instance +
        private static CCAReportDAL _instance;

        public static CCAReportDAL Instance
        {
            get
            {
                _instance = new CCAReportDAL();
                return _instance;
            }

        }
        #endregion

        #region GetDetailPerformanceReport 
        public DataTable GetDetailPerformanceReport(string startdate, string enddate)
        {
            try
            {
                cmd = new SqlCommand();
                //cmd.CommandTimeout = 200;
                cmd.CommandText = "[AACCIVR].[dbo].[sp_Detail_PerformanceCCA_Report_BAAC]";
                cmd.CommandType = System.Data.CommandType.StoredProcedure;

                #region Parameter 

                SqlParameter StartDate = new SqlParameter();
                StartDate.ParameterName = "@StartDate";
                StartDate.SqlDbType = SqlDbType.NVarChar;
                StartDate.Value = startdate;

                SqlParameter EndDate = new SqlParameter();
                EndDate.ParameterName = "@EndDate";
                EndDate.SqlDbType = SqlDbType.NVarChar;
                EndDate.Value = enddate;


                cmd.Parameters.Add(StartDate);
                cmd.Parameters.Add(EndDate);
                #endregion

                DataTable dt = dbManager.ExecuteDataTable(cmd, "DetailSelect");

                return dt;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        #endregion



        #region GetSummaryPerformanceReport 
        public DataTable GetSummaryPerformanceReport(string startdate, string enddate)
        {
            try
            {
                cmd = new SqlCommand();
                cmd.CommandTimeout = 1000;
                cmd.CommandText = "[AACCIVR].[dbo].[sp_Summary_PerformanceCCA_Report_BAAC]";
                cmd.CommandType = System.Data.CommandType.StoredProcedure;

                #region Parameter 


                SqlParameter StartDate = new SqlParameter();
                StartDate.ParameterName = "@startdate";
                StartDate.SqlDbType = SqlDbType.NVarChar;
                StartDate.Value = startdate;

                SqlParameter EndDate = new SqlParameter();
                EndDate.ParameterName = "@enddate";
                EndDate.SqlDbType = SqlDbType.NVarChar;
                EndDate.Value = enddate;

                cmd.Parameters.Add(StartDate);
                cmd.Parameters.Add(EndDate);
                #endregion

                DataTable dt = dbManager.ExecuteDataTable(cmd, "SUMMARYSELECT");

                return dt;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        #endregion

        /// <summary>
        /// เพิ่มเข้ามา
        /// </summary>
        /// <param name="startdate"></param>
        /// <param name="enddate"></param>
        /// <returns></returns>

        #region GetCaseDetail 
        public DataTable GetCaseDetail(string startdate, string enddate)
        {
            try
            {
                cmd = new SqlCommand();
                //cmd.CommandTimeout = 200;
                cmd.CommandText = "[BAAC].[dbo].[sp_CaseDetail_Report_BAAC]";
                cmd.CommandType = System.Data.CommandType.StoredProcedure;

                #region Parameter 

                SqlParameter StartDate = new SqlParameter();
                StartDate.ParameterName = "@StartDate";
                StartDate.SqlDbType = SqlDbType.NVarChar;
                StartDate.Value = startdate;

                SqlParameter EndDate = new SqlParameter();
                EndDate.ParameterName = "@EndDate";
                EndDate.SqlDbType = SqlDbType.NVarChar;
                EndDate.Value = enddate;


                cmd.Parameters.Add(StartDate);
                cmd.Parameters.Add(EndDate);
                #endregion

                DataTable dt = dbManager.ExecuteDataTable(cmd, "DetailSelect");

                return dt;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        #endregion

        #region GetOutboundDetail

        public DataTable GetOutboundDetail(string startdate, string enddate)
        {
            try
            {
                cmd = new SqlCommand();
                //cmd.CommandTimeout = 200;
                cmd.CommandText = "[BAAC].[dbo].[outbond_report]";
                cmd.CommandType = System.Data.CommandType.StoredProcedure;

                #region Parameter 

                SqlParameter StartDate = new SqlParameter();
                StartDate.ParameterName = "@StartDate";
                StartDate.SqlDbType = SqlDbType.NVarChar;
                StartDate.Value = startdate;

                SqlParameter EndDate = new SqlParameter();
                EndDate.ParameterName = "@EndDate";
                EndDate.SqlDbType = SqlDbType.NVarChar;
                EndDate.Value = enddate;


                cmd.Parameters.Add(StartDate);
                cmd.Parameters.Add(EndDate);
                #endregion

                DataTable dt = dbManager.ExecuteDataTable(cmd, "DetailSelect");

                return dt;
            }
            catch (Exception ex)
            {
                throw ex;
            }


        }
        #endregion
    }

}
