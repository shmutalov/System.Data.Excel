using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace System.Data.Excel.Models
{
    /// <summary>
    /// Represents Excel table
    /// </summary>
    internal class ExcelTable
    {
        public ExcelTable()
        {
            
        }

        public ExcelTable(string name)
        {
            Name = name;
        }

        /// <summary>
        /// Table name
        /// </summary>
        public string Name { get; set; }

        public List<ExcelColumn> Columns => new List<ExcelColumn>();
    }
}
