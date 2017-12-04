using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelDatabaseBridge {
    public static class Database {
        #region 屬性
        public static SqlDataAdapter adapter;
        public static SqlCommandBuilder commandBuilder;
        public static SqlCommand command;
        public static SqlConnection conn;
        public static SqlTransaction transaction;
        #endregion

        public static void DisposeCommand() {
            if (command.IsValid() == true) {
                command.Dispose();
            }

            if (conn.IsValid() == true) {
                conn.Dispose();
            }
        }

        public static void InitialCommand() {
            conn = new SqlConnection(Properties.Settings.Default.DatabaseConnectionString);
            conn.Open();
            command = new SqlCommand();
            command.Connection = conn;
        }

        public static void InitialTransaction() {
            if (conn.IsValid() == true && command.IsValid() == true) {
                transaction = conn.BeginTransaction();
                command.Transaction = transaction;
            }
        }

        public static DataTable GetDataTableFromSQL(string sqlStatement) {
            command.CommandText = sqlStatement;
            SqlDataReader reader = command.ExecuteReader();
            DataTable dataTable = new DataTable();
            dataTable.Load(reader);
            reader.Dispose();

            return dataTable;
        }
    }
}