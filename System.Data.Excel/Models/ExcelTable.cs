#region Copyright

/*
	Copyright (c) Sherzod Mutalov, 2015
	mailto:shmutalov@gmail.com
*/

#endregion

using System.Collections.Generic;
using System.ComponentModel;

namespace System.Data.Excel.Models
{
    /// <summary>
    /// Represents Excel table
    /// </summary>
    [DefaultProperty("Name")]
    internal class ExcelTable
    {
        public ExcelTable()
        {
            Columns = new List<ExcelColumn>();
        }

        public ExcelTable(string name)
            :this()
        {
            Name = name;
        }

        /// <summary>
        /// Table name
        /// </summary>
        public string Name { get; set; }

        public List<ExcelColumn> Columns { get; }
    }
}
