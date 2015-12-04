using System.Data.Excel.Enums;
using Excel;

namespace System.Data.Excel.Extensions
{
    /// <summary>
    /// Excel data reader extension methods
    /// </summary>
    internal static class ReaderExt
    {
        /// <summary>
        /// Resets reader cursor position
        /// </summary>
        /// <param name="reader"></param>
        /// <param name="type"></param>
        public static void Reset(this IExcelDataReader reader)
        {
            if (reader is ExcelBinaryReader)
            {
                reader.SetInstanceFieldValue("m_IsFirstRead", true);
            }
            else if (reader is ExcelOpenXmlReader)
            {
                reader.SetInstanceFieldValue("_isFirstReader", true);
            }
        }

        /// <summary>
        /// Shifts reader result set to specific table
        /// </summary>
        /// <param name="reader"></param>
        /// <param name="type"></param>
        /// <param name="resultSetId"></param>
        public static void SetTable(this IExcelDataReader reader, ExcelDocumentType type, int resultSetId)
        {
            if (reader is ExcelBinaryReader)
            {
                reader.SetInstanceFieldValue("m_SheetIndex", resultSetId);
            }
            else if (reader is ExcelOpenXmlReader)
            {
                reader.SetInstanceFieldValue("_resultIndex", resultSetId);
            }

            reader.Reset();
        }
    }
}
