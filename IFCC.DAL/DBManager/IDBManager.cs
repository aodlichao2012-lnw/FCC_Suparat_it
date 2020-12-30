using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;

namespace IFCC.DAL
{
    public enum DataProvider
    {
        SqlServer,
        //SqlServerCe,
        Oracle,
        OleDb,
        Odbc
    }

    public interface IDBManager
    {
        DataProvider ProviderType
        {
            get;
            set;
        }

        string ConnectionString
        {
            get;
            set;
        }

        IDbConnection Connection
        {
            get;
        }

        IDbTransaction Transaction
        {
            get;
        }

        IDataReader DataReader
        {
            get;
        }

        IDbCommand Command
        {
            get;
        }

        IDbDataParameter[] Parameters
        {
            get;
        }

        void Open();

        void BeginTransaction();

        void CommitTransaction();

        void RollbackTransaction();

        void CreateParameters(int paramsCount);

        void AddParameters(int index, string paramName, object objValue);

        IDataReader ExecuteReader(CommandType commandType, string commandText);

        DataSet ExecuteDataSet(CommandType commandType, string commandText);

        DataTable ExecuteDataTable(CommandType commandType, string commandText, string tableName);

        object ExecuteScalar(CommandType commandType, string commandText);

        int ExecuteNonQuery(CommandType commandType, string commandText);

        void CloseReader();

        void Close();

        void Dispose();

        IDataReader ExecuteReader(IDbCommand command);

        DataSet ExecuteDataSet(IDbCommand command);

        DataTable ExecuteDataTable(IDbCommand command, string tableName);

        object ExecuteScalar(IDbCommand command);

        int ExecuteNonQuery(IDbCommand command);
    }
}
