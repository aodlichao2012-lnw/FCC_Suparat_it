using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IFCC.DAL
{
    public class MasterDAL : BaseDAL
    {

        #region + Instance +
        private static MasterDAL _instance;

        public static MasterDAL Instance
        {
            get
            {
                _instance = new MasterDAL();
                return _instance;
            }

        }
        #endregion


        #region GetCallBack 
        public DataTable GetProjectName()
        {
            try
            {

                cmd = new SqlCommand();


                cmd.CommandText = "[AACCIVR].[dbo].[sp_Master_ProjectnameForCallBack]";
                cmd.CommandType = System.Data.CommandType.StoredProcedure;

                DataTable dt = dbManager.ExecuteDataTable(cmd, "MASTERPROJECTNAME");

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
