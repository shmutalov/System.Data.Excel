using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SqlServer.Server;

namespace System.Data.Excel.Models
{
    /// <summary>
    /// Represents Excel table column
    /// </summary>
    internal class ExcelColumn
    {
        public ExcelColumn()
        {
            DataType = typeof(string);
        }

        public ExcelColumn(string name)
        {
            Name = name;
        }

        public ExcelColumn(ExcelTable table, string name)
        {
            Table = table;
            Name = name;
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
