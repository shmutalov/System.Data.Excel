using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using JetBrains.Annotations;

namespace System.Data.Excel.Models
{
    /// <summary>
    /// Excel exception
    /// </summary>
    public class ExcelException : Exception
    {
        public ExcelException(string message, Exception innerException)
            : base(message, innerException)
        {

        }

        public ExcelException(string message)
            : base(message)
        {

        }

        [StringFormatMethod("format")]
        public ExcelException(string format, params object[] args)
            : base(string.Format(format, args))
        {

        }

        [StringFormatMethod("format")]
        public ExcelException(Exception innerException, string format,  params object[] args)
            : base(string.Format(format, args), innerException)
        {

        }
    }
}
