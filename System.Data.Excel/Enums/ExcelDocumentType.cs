﻿#region Copyright

/*
	Copyright (c) Sherzod Mutalov, 2015
	mailto:shmutalov@gmail.com
*/

#endregion

namespace System.Data.Excel.Enums
{
    /// <summary>
    /// Excel file types
    /// </summary>
    internal enum ExcelDocumentType
    {
        /// <summary>
        /// xlsx, Excel &gt; 2007
        /// </summary>
        Xml,

        /// <summary>
        /// xls, Excel &lt; 2007
        /// </summary>
        Binary,
    }
}
