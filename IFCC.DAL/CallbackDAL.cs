using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.Data;


namespace IFCC.DAL
{
   public class CallbackDAL : BaseDAL
    {

        #region + Instance +
        private static CallbackDAL _instance;

        public static CallbackDAL Instance
        {
            get
            {
                _instance = new CallbackDAL();
                return _instance;
            }

        }
        #endregion

        #region GetCallBack 
        public DataTable GetCallBack(string data, string startdate,string enddate)
        {
            try
            {
                cmd = new SqlCommand();
                cmd.CommandText = "[AACCIVR].[dbo].[sp_GetCallBackSec]";
                cmd.CommandType = System.Data.CommandType.StoredProcedure;

                #region Parameter 
                SqlParameter Param1 = new SqlParameter();
                Param1.ParameterName = "@projectname";
                Param1.SqlDbType = SqlDbType.NVarChar;
                Param1.Direction = ParameterDirection.Input;
                Param1.Value = data;
                SqlParameter StartDate = new SqlParameter();
                StartDate.ParameterName = "@StartDate";
                StartDate.SqlDbType = SqlDbType.DateTime;
                StartDate.Value = startdate;

                SqlParameter EndDate = new SqlParameter();
                EndDate.ParameterName = "@EndDate";
                EndDate.SqlDbType = SqlDbType.DateTime;
                EndDate.Value = enddate;

                cmd.Parameters.Add(Param1);
                cmd.Parameters.Add(StartDate);
                cmd.Parameters.Add(EndDate);
                #endregion

                DataTable dt = dbManager.ExecuteDataTable(cmd, "CALLBACKSELECT");

                return dt;
            } catch (Exception ex) 
            { 
                throw ex; 
            }
        }
        #endregion


        #region GetUpdateStatus 
        public DataTable GetUpdateStatus(string id, string status)
        {
            try
            {
                cmd = new SqlCommand();
                cmd.CommandText = "[AACCIVR].[dbo].[sp_GetUpdateStatusCallBack]";
                cmd.CommandType = System.Data.CommandType.StoredProcedure;

                #region Parameter 
                SqlParameter ID = new SqlParameter();
                ID.ParameterName = "@id";
                ID.SqlDbType = SqlDbType.NVarChar;
                ID.Direction = ParameterDirection.Input;
                ID.Value = id;

                SqlParameter STATUS = new SqlParameter();
                STATUS.ParameterName = "@status";
                STATUS.SqlDbType = SqlDbType.NVarChar;
                STATUS.Direction = ParameterDirection.Input;
                STATUS.Value = status;


                cmd.Parameters.Add(ID);
                cmd.Parameters.Add(STATUS);

                #endregion
                dbManager.Open();
                dbManager.BeginTransaction();
                dbManager.ExecuteNonQuery(cmd);
                dbManager.CommitTransaction();
                SetReturnData();
                DataTable dt = new DataTable();
                dt.Rows.Add("Sucess");
                _ = dt.Rows.Count;
                return  dt;

            }

            catch (Exception ex)
            {
                throw ex;
            }
        }
        #endregion
    }
}
