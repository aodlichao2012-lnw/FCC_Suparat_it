using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Odbc;
using System.Data.OleDb;
using System.Data.SqlClient;

//using System.Data.SqlServerCe;
using System.Linq;
using System.Text;

namespace IFCC.DAL
{
    public sealed class DBManagerFactory
    {
        private DBManagerFactory()
        {
        }

        public static IDbConnection GetConnection(DataProvider providerType)
        {
            IDbConnection dbConnection;
            IDbConnection result;
            try
            {
                switch (providerType)
                {
                    case DataProvider.SqlServer:
                        dbConnection = new SqlConnection();
                        break;
                    //case DataProvider.SqlServerCe:
                    //	dbConnection = new SqlCeConnection();
                    //	break;
                    //case DataProvider.Oracle:
                    //    dbConnection = new OracleConnection();
                    //    break;
                    case DataProvider.OleDb:
                        dbConnection = new OleDbConnection();
                        break;
                    case DataProvider.Odbc:
                        dbConnection = new OdbcConnection();
                        break;
                    default:
                        result = null;
                        return result;
                }
            }
            catch (Exception)
            {
                dbConnection = new SqlConnection();
            }
            result = dbConnection;
            return result;
        }

        public static IDbCommand GetCommand(DataProvider providerType)
        {
            IDbCommand result;
            switch (providerType)
            {
                case DataProvider.SqlServer:
                    result = new SqlCommand();
                    break;
                //case DataProvider.SqlServerCe:
                //	result = new SqlCeCommand();
                //	break;
                //case DataProvider.Oracle:
                //    result = new OracleCommand();
                //    break;
                case DataProvider.OleDb:
                    result = new OleDbCommand();
                    break;
                case DataProvider.Odbc:
                    result = new OdbcCommand();
                    break;
                default:
                    result = null;
                    break;
            }
            return result;
        }

        public static IDbDataAdapter GetDataAdapter(DataProvider providerType)
        {
            IDbDataAdapter result;
            switch (providerType)
            {
                case DataProvider.SqlServer:
                    result = new SqlDataAdapter();
                    break;
                //case DataProvider.SqlServerCe:
                //	result = new SqlCeDataAdapter();
                //	break;
                //case DataProvider.Oracle:
                //    result = new OracleDataAdapter();
                //    break;
                case DataProvider.OleDb:
                    result = new OleDbDataAdapter();
                    break;
                case DataProvider.Odbc:
                    result = new OdbcDataAdapter();
                    break;
                default:
                    result = null;
                    break;
            }
            return result;
        }

        public static IDbTransaction GetTransaction(DataProvider providerType)
        {
            IDbConnection connection = DBManagerFactory.GetConnection(providerType);
            return connection.BeginTransaction();
        }

        public static IDataParameter GetParameter(DataProvider providerType)
        {
            IDataParameter result = null;
            switch (providerType)
            {
                case DataProvider.SqlServer:
                    result = new SqlParameter();
                    break;
                //case DataProvider.SqlServerCe:
                //	result = new SqlCeParameter();
                //	break;
                //case DataProvider.Oracle:
                //    result = new OracleParameter();
                //    break;
                case DataProvider.OleDb:
                    result = new OleDbParameter();
                    break;
                case DataProvider.Odbc:
                    result = new OdbcParameter();
                    break;
            }
            return result;
        }

        public static IDbDataParameter[] GetParameters(DataProvider providerType, int paramsCount)
        {
            IDbDataParameter[] array = new IDbDataParameter[paramsCount];
            switch (providerType)
            {
                case DataProvider.SqlServer:
                    for (int i = 0; i < paramsCount; i++)
                    {
                        array[i] = new SqlParameter();
                    }
                    break;
                //case DataProvider.SqlServerCe:
                //	for(int i = 0; i < paramsCount; i++)
                //	{
                //		array[i] = new SqlCeParameter();
                //	}
                //	break;
                //case DataProvider.Oracle:
                //    for (int i = 0; i < paramsCount; i++)
                //    {
                //        array[i] = new OracleParameter();
                //    }
                //    break;
                case DataProvider.OleDb:
                    for (int i = 0; i < paramsCount; i++)
                    {
                        array[i] = new OleDbParameter();
                    }
                    break;
                case DataProvider.Odbc:
                    for (int i = 0; i < paramsCount; i++)
                    {
                        array[i] = new OdbcParameter();
                    }
                    break;
                default:
                    array = null;
                    break;
            }
            return array;
        }
    }
}
