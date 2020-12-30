using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
//using System.Data.OracleClient;
using System.Data.SqlClient;

namespace IFCC.DAL
{
    public class BaseDAL
    {
        protected SqlCommand cmd = null;
        protected int IsSuccess = 0;
        protected string MsgErrorRowCount = "ROWCOUNT equal 0";
        protected DataTable dtValue;
        protected Dictionary<string, object> dicResult = new Dictionary<string, object>();
        protected DBManager dbManager = new DBManager(DataProvider.SqlServer, ConfigurationManager.ConnectionStrings["sqlConnection"].ConnectionString);
        protected string connectionString = ConfigurationManager.ConnectionStrings["sqlConnection"].ConnectionString;

        //protected BaseDAL()
        //{
        //    if (dbManager != null)
        //    {
        //        dbManager.Close();
        //    }
        //    dbManager = new DBManager(DataProvider.SqlServer, ConfigurationManager.ConnectionStrings["sqlConnection"].ConnectionString);
        //    if (dbManager.Connection.State == ConnectionState.Closed)
        //    {
        //        dbManager.Open();
        //    }
        //}

        #region CheckROWCOUNT
        protected void CheckROWCOUNT()
        {
            if (dtValue != null && dtValue.Rows.Count > 0)
            {
                if (dtValue.Rows[0]["ROWCOUNT"] + string.Empty == "0")
                {
                    throw new Exception(MsgErrorRowCount);
                }
            }
        }
        #endregion

        #region Set return data
        protected void SetReturnData()
        {
            dicResult.Add("status", 0);
            dicResult.Add("return_data", dtValue);
        }
        #endregion

        #region Set return status
        protected void SetReturnStatus(int status)
        {
            dicResult.Add("status", status);
        }
        #endregion

        #region Set return status
        protected void SetPDFReturnStatus(string refData,string refVendor)
        {
            dicResult.Add("status", refData);
            dicResult.Add("status2", refVendor);
        }
        #endregion

        #region SetTableName
        protected void SetTableName(ref DataSet ds)
        {
            if (ds != null && ds.Tables != null)
            {
                for (int i = 0; i < ds.Tables.Count; i++)
                {
                    if (ds.Tables[i].Rows.Count > 0)
                    {
                        ds.Tables[i].TableName = ds.Tables[i].Rows[0]["TableName"] + string.Empty;
                    }
                }
            }
            
        }
        #endregion
    }
}
