#region Copyright

/*
	Copyright (c) Sherzod Mutalov, 2015
	mailto:shmutalov@gmail.com
*/

#endregion

using System.ComponentModel;

namespace System.Data.Excel.Models
{
    /// <summary>
    /// Represents Excel table column
    /// </summary>
    [DefaultProperty("Name")]
    internal class ExcelColumn
    {
        public ExcelColumn()
        {
            DataType = typeof(string);
        }

        public ExcelColumn(string name)
            : this()
        {
            Name = name;
        }

        public ExcelColumn(ExcelTable table, string name)
            : this(name)
        {
            Table = table;
        }

        /// <summary>
        /// Column name
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// Column data type
        /// </summary>
        public Type DataType { get; set; }

        /// <summary>
        /// Column's parent table
        /// </summary>
        public ExcelTable Table { get; set; }
    }
}
