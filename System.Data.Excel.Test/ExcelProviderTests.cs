using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NUnit.Framework;

namespace System.Data.Excel.Test
{
    [TestFixture]
    public class ExcelProviderTests
    {
        private const string SmallExcelFile = @"C:\Superstore_Locked.xlsx";
        private const string LargeExcelFile = @"C:\AWLArge_FactInternetSales_1M.xlsx";

        private const string ConnectionStringTemplate = "Database={0};Type=Xml;Password=;FirstRowIsHeader=True;StorageDir=;";

        [Test]
        public void TestExcelConnectionToSmallFile()
        {
            using (var connection = new ExcelConnection(string.Format(ConnectionStringTemplate, SmallExcelFile)))
            {
                connection.Open();

                connection.Close();
            }
        }
    }
}
