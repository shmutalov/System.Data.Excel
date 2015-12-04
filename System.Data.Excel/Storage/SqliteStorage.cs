using System.Collections.Generic;
using System.Data.Excel.Helpers;
using System.Data.Excel.Models;
using System.Data.SQLite;
using System.IO;
using System.Text;
using Excel;

namespace System.Data.Excel.Storage
{
    /// <summary>
    /// SQLite storage
    /// </summary>
    internal class SqliteStorage : IStorage
    {
        private const string DatabaseNameTemplate = "{0}.db3";
        private const string ConnectionStringTemplate = "Data Source={0};Version=3;";

        public string GetConnectionString(string database, string user, string password)
        {
            return string.Format(ConnectionStringTemplate, database);
        }

        public IDbConnection GetConnection(string connectionString)
        {
            return new SQLiteConnection(connectionString);
        }

        public string GetDatabaseName(string excelFileName, string storageDir)
        {
            return Path.Combine(
                storageDir,
                string.Format(
                    DatabaseNameTemplate,
                    Path.GetFileName(excelFileName)));
        }

        public void CreateDatabase(string database, string storageDir)
        {
            SQLiteConnection.CreateFile(Path.Combine(storageDir, database));
        }

        public void DropDatabase(string database, string storageDir)
        {
            try
            {
                File.Delete(Path.Combine(storageDir, database));
            }
            catch (Exception)
            {
                // ignore
            }
        }

        public bool DatabaseExists(string database, string storageDir)
        {
            return File.Exists(Path.Combine(storageDir, database));
        }

        public void ImportData(IExcelDataReader sourceReader, bool firstRowIsHeader, IDbConnection storageConnection)
        {
            do
            {
                List<object[]> preloadedValues;
                var table = ExcelHelper.GetTable(sourceReader, firstRowIsHeader, out preloadedValues);

                CreateTable((SQLiteConnection)storageConnection, table);

                // upload preloaded values

                while (sourceReader.Read())
                {
                    // ignore
                }
            } while (sourceReader.NextResult());
        }

        public string GetStorageDataType(Type type)
        {
            string result;

            switch (Type.GetTypeCode(type))
            {
                case TypeCode.Boolean:
                case TypeCode.SByte:
                case TypeCode.Byte:
                case TypeCode.Int16:
                case TypeCode.UInt16:
                    result = "INTEGER";
                    break;
                case TypeCode.Int32:
                case TypeCode.UInt32:
                case TypeCode.Int64:
                case TypeCode.UInt64:
                    result = "BIGINT";
                    break;
                case TypeCode.Single:
                case TypeCode.Double:
                case TypeCode.Decimal:
                    result = "DOUBLE";
                    break;
                case TypeCode.DateTime:
                    result = "DATETIME";
                    break;
                default:
                    result = "TEXT";
                    break;
            }

            return result;
        }

        /// <summary>
        /// Create storage table form excel table
        /// </summary>
        /// <param name="conenction"></param>
        /// <param name="table"></param>
        private void CreateTable(SQLiteConnection conenction, ExcelTable table)
        {
            try
            {
                var sb = new StringBuilder();

                sb.AppendLine(string.Format("CREATE TABLE IF NOT EXISTS `{0}`", table.Name));
                sb.AppendLine("(");

                // columns
                for (var columnId = 0; columnId < table.Columns.Count; columnId++)
                {
                    if (columnId > 0)
                        sb.AppendLine(",");

                    var column = table.Columns[columnId];

                    var storageDataType = GetStorageDataType(column.DataType);

                    // NOTE: Excel columns always nullable
                    sb.AppendFormat("\t`{0}` {1} NULL", column.Name, storageDataType);
                }

                sb.AppendLine();
                sb.AppendLine(")");

                using (var cmd = conenction.CreateCommand())
                {
                    cmd.CommandText = sb.ToString();
                    cmd.ExecuteNonQuery();
                }
            }
            catch (Exception ex)
            {
                throw new ExcelException(ex, "Cannot create storage table '{0}'", table.Name);
            }
        }
    }
}
