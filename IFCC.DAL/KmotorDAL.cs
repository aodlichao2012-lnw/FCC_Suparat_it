using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.Data;


namespace IFCC.DAL
{
   public class KmotorDAL : BaseDAL
    {

        #region + Instance +
        private static KmotorDAL _instance;

        public static KmotorDAL Instance
        {
            get
            {
                _instance = new KmotorDAL();
                return _instance;
            }

        }
        #endregion

        #region GetKmotorReport
        public DataTable GetKmotorReport(string startdate)
        {
            try
            {
                cmd = new SqlCommand();
                cmd.CommandText = "[AACCIVR].[dbo].[sp_Kmortor_Report]";
                cmd.CommandType = System.Data.CommandType.StoredProcedure;

                #region Parameter 
      
                SqlParameter StartDate = new SqlParameter();
                StartDate.ParameterName = "@startdate";
                StartDate.SqlDbType = SqlDbType.NVarChar;
                StartDate.Value = startdate;

                cmd.Parameters.Add(StartDate);
              
                #endregion

                DataTable dt = dbManager.ExecuteDataTable(cmd, "CALLBACKSELECT");

                return dt;
            } catch (Exception ex) 
            { 
                throw ex; 
            }
        }
        #endregion


       
    }
}
