using System.Data.Excel.Constants;
using System.Data.Excel.Enums;
using System.Data.Excel.Extensions;

namespace System.Data.Excel.Models
{
    /// <summary>
    /// Excel connection parameters
    /// </summary>
    internal class ExcelConnectionParameters
    {
        /// <summary>
        /// Database name (document name)
        /// </summary>
        public string Database { get; set; }

        /// <summary>
        /// Storage directory
        /// </summary>
        public string StoregeDirectory { get; set; }

        /// <summary>
        /// Database type (file type)
        /// </summary>
        public ExcelDocumentType Type { get; set; }

        /// <summary>
        /// Password to database
        /// </summary>
        public string Password { get; set; }

        /// <summary>
        /// First row of tables are table column names
        /// </summary>
        public bool FirstRowIsHeader { get; set; }

        /// <summary>
        /// Forces reload all data of internal storage
        /// </summary>
        public bool ForceStorageReload { get; set; }

        /// <summary>
        /// Build connection parameters by parsing connection string
        /// </summary>
        /// <param name="connectionString">Connection string to parse</param>
        /// <returns></returns>
        public static ExcelConnectionParameters FromConnectionString(string connectionString)
        {
            var parameters = new ExcelConnectionParameters();

            var splitted = connectionString.Split(new[] {";"}, StringSplitOptions.RemoveEmptyEntries);

            if (splitted.Length == 0)
                return parameters;

            foreach (var entry in splitted)
            {
                var splittedKeyVal = entry.Split(new[] {"="}, StringSplitOptions.RemoveEmptyEntries);

                if (splittedKeyVal.Length == 2)
                {
                    switch (splittedKeyVal[0].ToUpper())
                    {
                        case ExcelConnectionParameterNames.Database:
                            parameters.Database = splittedKeyVal[1];
                            break;
                        case ExcelConnectionParameterNames.StorageDirectory:
                            parameters.Database = splittedKeyVal[1];
                            break;
                        case ExcelConnectionParameterNames.Password:
                            parameters.Password = splittedKeyVal[1];
                            break;
                        case ExcelConnectionParameterNames.Type:
                        {
                            ExcelDocumentType documentType;
                            Enum.TryParse(splittedKeyVal[1], true, out documentType);
                            parameters.Type = documentType;
                        }
                            break;
                        case ExcelConnectionParameterNames.FirstRowIsHeader:
                            parameters.FirstRowIsHeader = splittedKeyVal[1].ToBool();
                            break;
                        case ExcelConnectionParameterNames.ForceStorageReload:
                            parameters.ForceStorageReload = splittedKeyVal[1].ToBool();
                            break;
                        default:
                            continue;
                    }
                }
            }

            return parameters;
        }

        public static string ToConnectionString(ExcelConnectionParameters parameters)
        {
            return string.Format(
                "{0}={1};{2}={3};{4}={5};{6}={7};{8}={9}",
                ExcelConnectionParameterNames.Database, parameters.Database,
                ExcelConnectionParameterNames.StorageDirectory, parameters.StoregeDirectory,
                ExcelConnectionParameterNames.Password, parameters.Password,
                ExcelConnectionParameterNames.Type, parameters.Type,
                ExcelConnectionParameterNames.FirstRowIsHeader, parameters.FirstRowIsHeader);
        }
    }
}