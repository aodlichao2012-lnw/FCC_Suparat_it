using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;

namespace IFCC.DAL
{
    public sealed class DBManager : IDBManager, IDisposable
    {
        private IDbConnection idbConnection;

        private IDataReader idataReader;

        private IDbCommand idbCommand;

        private DataProvider providerType;

        private IDbTransaction idbTransaction = null;

        private IDbDataParameter[] idbParameters = null;

        private string strConnection;

        public IDbConnection Connection
        {
            get
            {
                return this.idbConnection;
            }
        }

        public IDataReader DataReader
        {
            get
            {
                return this.idataReader;
            }
            set
            {
                this.idataReader = value;
            }
        }

        public DataProvider ProviderType
        {
            get
            {
                return this.providerType;
            }
            set
            {
                this.providerType = value;
            }
        }

        public string ConnectionString
        {
            get
            {
                return this.strConnection;
            }
            set
            {
                this.strConnection = value;
            }
        }

        public IDbCommand Command
        {
            get
            {
                return this.idbCommand;
            }
        }

        public IDbTransaction Transaction
        {
            get
            {
                return this.idbTransaction;
            }
        }

        public IDbDataParameter[] Parameters
        {
            get
            {
                return this.idbParameters;
            }
        }

        public DBManager()
        {
        }

        public DBManager(DataProvider providerType)
        {
            this.providerType = providerType;
        }

        public DBManager(DataProvider providerType, string connectionString)
        {
            this.providerType = providerType;
            this.strConnection = connectionString;
            this.idbConnection = DBManagerFactory.GetConnection(this.providerType);
            this.idbConnection.ConnectionString = connectionString;
        }

        public void Open()
        {
            if (this.idbConnection == null)
            {
                this.idbConnection = DBManagerFactory.GetConnection(this.providerType);
                this.idbConnection.ConnectionString = this.ConnectionString;
            }
            if (this.idbConnection.State != ConnectionState.Open)
            {
                this.idbConnection.Open();
            }
            this.idbCommand = DBManagerFactory.GetCommand(this.ProviderType);
        }

        public void Close()
        {
            if (this.idbConnection.State != ConnectionState.Closed)
            {
                this.idbConnection.Close();
            }
        }

        public void Dispose()
        {
            GC.SuppressFinalize(this);
            this.Close();
            this.idbCommand = null;
            this.idbTransaction = null;
            this.idbConnection = null;
        }

        public void CreateParameters(int paramsCount)
        {
            this.idbParameters = new IDbDataParameter[paramsCount];
            this.idbParameters = DBManagerFactory.GetParameters(this.ProviderType, paramsCount);
        }

        public void AddParameters(int index, string paramName, object objValue)
        {
            if (index < this.idbParameters.Length)
            {
                this.idbParameters[index].ParameterName = paramName;
                this.idbParameters[index].Value = objValue;
            }
        }

        public void BeginTransaction()
        {
            if (this.idbTransaction == null)
            {
                this.idbTransaction = this.Connection.BeginTransaction();
            }
            this.idbCommand.Transaction = this.idbTransaction;
        }

        public void CommitTransaction()
        {
            if (this.idbTransaction != null)
            {
                this.idbTransaction.Commit();
            }
            this.idbTransaction = null;
        }

        public void RollbackTransaction()
        {
            if (this.idbTransaction != null)
            {
                this.idbTransaction.Rollback();
            }
            this.idbTransaction = null;
        }

        public IDataReader ExecuteReader(CommandType commandType, string commandText)
        {
            this.idbCommand = DBManagerFactory.GetCommand(this.ProviderType);
            this.idbCommand.Connection = this.Connection;
            this.PrepareCommand(this.idbCommand, this.Connection, this.Transaction, commandType, commandText, this.Parameters);
            this.DataReader = this.idbCommand.ExecuteReader();
            this.idbCommand.Parameters.Clear();
            return this.DataReader;
        }

        public void CloseReader()
        {
            if (this.DataReader != null)
            {
                this.DataReader.Close();
            }
        }

        private void AttachParameters(IDbCommand command, IDbDataParameter[] commandParameters)
        {
            for (int i = 0; i < commandParameters.Length; i++)
            {
                IDbDataParameter dbDataParameter = commandParameters[i];
                if (dbDataParameter.Direction == ParameterDirection.InputOutput && dbDataParameter.Value == null)
                {
                    dbDataParameter.Value = DBNull.Value;
                }
                command.Parameters.Add(dbDataParameter);
            }
        }

        private void PrepareCommand(IDbCommand command, IDbConnection connection, IDbTransaction transaction, CommandType commandType, string commandText, IDbDataParameter[] commandParameters)
        {
            command.Connection = connection;
            command.CommandText = commandText;
            command.CommandType = commandType;
            if (transaction != null)
            {
                command.Transaction = transaction;
            }
            if (commandParameters != null)
            {
                this.AttachParameters(command, commandParameters);
            }
        }

        public int ExecuteNonQuery(CommandType commandType, string commandText)
        {
            this.idbCommand = DBManagerFactory.GetCommand(this.ProviderType);
            this.PrepareCommand(this.idbCommand, this.Connection, this.Transaction, commandType, commandText, this.Parameters);
            int result = this.idbCommand.ExecuteNonQuery();
            this.idbCommand.Parameters.Clear();
            return result;
        }

        public object ExecuteScalar(CommandType commandType, string commandText)
        {
            this.idbCommand = DBManagerFactory.GetCommand(this.ProviderType);
            this.PrepareCommand(this.idbCommand, this.Connection, this.Transaction, commandType, commandText, this.Parameters);
            object result = this.idbCommand.ExecuteScalar();
            this.idbCommand.Parameters.Clear();
            return result;
        }

        public DataSet ExecuteDataSet(CommandType commandType, string commandText)
        {
            this.idbCommand = DBManagerFactory.GetCommand(this.ProviderType);
            this.PrepareCommand(this.idbCommand, this.Connection, this.Transaction, commandType, commandText, this.Parameters);
            IDbDataAdapter dataAdapter = DBManagerFactory.GetDataAdapter(this.ProviderType);
            dataAdapter.SelectCommand = this.idbCommand;
            DataSet dataSet = new DataSet();
            dataAdapter.Fill(dataSet);
            this.idbCommand.Parameters.Clear();
            return dataSet;
        }

        internal DataTable ExecuteDataTable(SqlCommand cmd)
        {
            throw new NotImplementedException();
        }

        public DataTable ExecuteDataTable(CommandType commandType, string commandText, string tableName)
        {
            this.idbCommand = DBManagerFactory.GetCommand(this.ProviderType);
            this.PrepareCommand(this.idbCommand, this.Connection, this.Transaction, commandType, commandText, this.Parameters);
            IDbDataAdapter dataAdapter = DBManagerFactory.GetDataAdapter(this.ProviderType);
            dataAdapter.SelectCommand = this.idbCommand;
            DataSet dataSet = new DataSet();
            dataAdapter.Fill(dataSet);
            DataTable dataTable = dataSet.Tables[0];
            dataTable.TableName = tableName;
            this.idbCommand.Parameters.Clear();
            return dataTable;
        }

        private void PrepareCommand(IDbCommand command, IDbConnection connection, IDbTransaction transaction)
        {
            command.Connection = connection;
            if (transaction != null)
            {
                command.Transaction = transaction;
            }
        }

        public IDataReader ExecuteReader(IDbCommand command)
        {
            this.idbCommand = command;
            this.idbCommand.Connection = this.Connection;
            this.PrepareCommand(this.idbCommand, this.Connection, this.Transaction);
            this.DataReader = this.idbCommand.ExecuteReader();
            return this.DataReader;
        }

        public int ExecuteNonQuery(IDbCommand command)
        {
            this.idbCommand = command;
            this.PrepareCommand(this.idbCommand, this.Connection, this.Transaction);
            return this.idbCommand.ExecuteNonQuery();
        }

        public DataSet ExecuteDataSet(IDbCommand command)
        {
            this.idbCommand = command;
            this.PrepareCommand(this.idbCommand, this.Connection, this.Transaction);
            IDbDataAdapter dataAdapter = DBManagerFactory.GetDataAdapter(this.ProviderType);
            dataAdapter.SelectCommand = this.idbCommand;
            DataSet dataSet = new DataSet();
            dataAdapter.Fill(dataSet);
            return dataSet;
        }

        public object ExecuteScalar(IDbCommand command)
        {
            this.idbCommand = command;
            this.PrepareCommand(this.idbCommand, this.Connection, this.Transaction);
            return this.idbCommand.ExecuteScalar();
        }

        public DataTable ExecuteDataTable(IDbCommand command, string tableName)
        {
            this.idbCommand = command;
            this.PrepareCommand(this.idbCommand, this.Connection, this.Transaction);
            IDbDataAdapter dataAdapter = DBManagerFactory.GetDataAdapter(this.ProviderType);
            dataAdapter.SelectCommand = this.idbCommand;
            DataTable dataTable = new DataTable(tableName);
            DataSet dataSet = new DataSet();
            dataAdapter.Fill(dataSet);
            dataTable = dataSet.Tables[0];
            dataTable.TableName = tableName;
            return dataTable;
        }
    }
}
