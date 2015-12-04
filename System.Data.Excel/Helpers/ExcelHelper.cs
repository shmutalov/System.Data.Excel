using System.Collections.Generic;
using System.Data.Excel.Extensions;
using System.Data.Excel.Models;
using System.Linq;
using Excel;
using JetBrains.Annotations;

namespace System.Data.Excel.Helpers
{
    internal static class ExcelHelper
    {
        /// <summary>
        /// Returns best data type for the given column by parsing rows
        /// </summary>
        /// <param name="dataTypes"></param>
        /// <param name="rows"></param>
        /// <param name="columnId"></param>
        /// <returns></returns>
        private static Type GetColumnDataType(object[][] dataTypes, int rows, int columnId)
        {
            var typesDict = new Dictionary<Type, int>();
            var result = typeof(string);

            if (columnId >= dataTypes.GetLength(1))
                return result;

            for (var rowId = 0; rowId < rows; rowId++)
            {
                var type = dataTypes[rowId][columnId].GetType();

                if (!typesDict.ContainsKey(type))
                    typesDict[type] = 0;

                typesDict[type]++;
            }

            result = typesDict.Aggregate((l, r) => l.Value > r.Value ? l : r).Key
                ?? typeof(string);

            return result;
        }

        /// <summary>
        /// Builds table from reader's result set
        /// </summary>
        /// <param name="reader"></param>
        /// <param name="firstRowIsHeader"></param>
        /// <param name="preloadedValues"></param>
        /// <returns></returns>
        public static ExcelTable GetTable([NotNull] IExcelDataReader reader, bool firstRowIsHeader, out List<object[]> preloadedValues)
        {
            reader.Reset();
            reader.Read();

            var table = new ExcelTable(reader.Name);
            var columnsCount = reader.FieldCount;

            if (firstRowIsHeader)
            {
                for (var columnId = 0; columnId < columnsCount; columnId++)
                {
                    var columnName = reader.GetString(columnId)
                                    ?? string.Format("Column {0}", columnId);

                    table.Columns.Add(new ExcelColumn(table, columnName));
                }
            }
            else
            {
                for (var columnId = 0; columnId < columnsCount; columnId++)
                {
                    table.Columns.Add(new ExcelColumn(table, string.Format("Column {0}", columnId)));
                }

                reader.Reset();
            }

            // preloaded data list
            var preloadedDataList = new List<object[]>();

            for (var rowId = 0; rowId < 10; rowId++)
            {
                if (!reader.Read())
                    break;

                var values = new object[columnsCount];

                for (var columnId = 0; columnId < columnsCount; columnId++)
                {
                    values[columnId] = reader.GetValue(columnId);
                }

                preloadedDataList.Add(values);
            }

            // calculate columns data types

            var data = preloadedDataList.ToArray();

            for (var columnId = 0; columnId < columnsCount; columnId++)
            {
                table.Columns[columnId].DataType = GetColumnDataType(data, preloadedDataList.Count, columnId);
            }

            preloadedValues = preloadedDataList;

            return table;
        }
    }
}
