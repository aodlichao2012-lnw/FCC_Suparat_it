using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.Data;


namespace IFCC.DAL
{
   public class BlacktieDAL : BaseDAL
    {

        #region + Instance +
        private static BlacktieDAL _instance;

        public static BlacktieDAL Instance
        {
            get
            {
                _instance = new BlacktieDAL();
                return _instance;
            }

        }
        #endregion

        #region GetCovid_19 
        public DataTable GetCovid_19(string startdate,string enddate)
        {
            try
            {
                cmd = new SqlCommand();
                cmd.CommandText = "[Blacktie].[dbo].[sp_Report_Covid-19]";
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

                DataTable dt = dbManager.ExecuteDataTable(cmd, "COVID_19SELECT");

                return dt;
            } catch (Exception ex) 
            { 
                throw ex; 
            }
        }
        #endregion




        #region GetGrabFood 
        public DataTable GetGrabFood(string startdate, string enddate)
        {
            try
            {
                cmd = new SqlCommand();
                cmd.CommandText = "[Blacktie].[dbo].[sp_Report-Grab_Food]";
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

                DataTable dt = dbManager.ExecuteDataTable(cmd, "GrabFoodSELECT");

                return dt;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        #endregion

        #region GetHealthCarePlus 
        public DataTable GetHealthCarePlus(string startdate, string enddate)
        {
            try
            {
                cmd = new SqlCommand();
                cmd.CommandText = "[Blacktie].[dbo].[sp_Report_Health_Care_Plus]";
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

                DataTable dt = dbManager.ExecuteDataTable(cmd, "HealthCareSELECT");

                return dt;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        #endregion

        #region GetHotel 
        public DataTable GetHotel(string startdate, string enddate)
        {
            try
            {
                cmd = new SqlCommand();
                cmd.CommandText = "[Blacktie].[dbo].[sp_Report_Hotel]";
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

                DataTable dt = dbManager.ExecuteDataTable(cmd, "HealthCareSELECT");

                return dt;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        #endregion

        #region GetRedCross
        public DataTable GetRedCross(string startdate, string enddate)
        {
            try
            {
                cmd = new SqlCommand();
                cmd.CommandText = "[Blacktie].[dbo].[sp_Report_RedCross]";
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

                DataTable dt = dbManager.ExecuteDataTable(cmd, "HealthCareSELECT");

                return dt;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        #endregion

        #region GetS_and_P
        public DataTable GetS_and_P(string startdate, string enddate)
        {
            try
            {
                cmd = new SqlCommand();
                cmd.CommandText = "[Blacktie].[dbo].[sp_Report_S_and_P]";
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

                DataTable dt = dbManager.ExecuteDataTable(cmd, "HealthCareSELECT");

                return dt;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        #endregion



        #region GetVaccine
        public DataTable GetVaccine(string startdate, string enddate)
        {
            try
            {
                cmd = new SqlCommand();
                cmd.CommandText = "[Blacktie].[dbo].[sp_Report_Vaccine]";
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

                DataTable dt = dbManager.ExecuteDataTable(cmd, "HealthCareSELECT");

                return dt;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        #endregion



        #region GetWine
        public DataTable GetWine(string startdate, string enddate)
        {
            try
            {
                cmd = new SqlCommand();
                cmd.CommandText = "[Blacktie].[dbo].[sp_Report_Wine]";
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

                DataTable dt = dbManager.ExecuteDataTable(cmd, "HealthCareSELECT");

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
