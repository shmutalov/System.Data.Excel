namespace System.Data.Excel.Constants
{
    /// <summary>
    /// Excel datasource connection string parameter names
    /// </summary>
    internal static class ExcelConnectionParameterNames
    {
        /// <summary>
        /// Database name (document file name)
        /// </summary>
        public const string Database = "DATABASE";

        /// <summary>
        /// Directory where internal storage must initialized
        /// </summary>
        public const string StorageDirectory = "STORAGEDIR";

        /// <summary>
        /// Password to decrypt document
        /// </summary>
        public const string Password = "PASSWORD";

        /// <summary>
        /// Excel file type
        /// </summary>
        public const string Type = "TYPE";

        /// <summary>
        /// First row of tables is header
        /// </summary>
        public const string FirstRowIsHeader = "FIRSTROWISHEADER";

        /// <summary>
        /// Forces internal storage to reinitialize
        /// </summary>
        public const string ForceStorageReload = "FORCESTORAGERELOAD";
    }
}
