using System;
using System.Collections.Generic;
using System.Diagnostics;
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

        private const string ConnectionStringTemplate = "Database={0};Type=Xml;Password=;FirstRowIsHeader=True;StorageDir=;ForceStorageReload=True";

        [Test]
        public void TestExcelConnectionToSmallFile()
        {
            var sw = new Stopwatch();

            sw.Start();

            using (var connection = new ExcelConnection(string.Format(ConnectionStringTemplate, SmallExcelFile)))
            {
                connection.Open();

                using (var cmd = connection.CreateCommand())
                {
                    cmd.CommandText = "SELECT 1";
                    var value = cmd.ExecuteScalar();

                    Assert.AreEqual(1, value);
                }

                connection.Close();
            }

            sw.Stop();

            Console.WriteLine("Test finished in {0} ms.", sw.ElapsedMilliseconds);
        }

        [Test]
        public void TestExcelConnectionToLargeFile()
        {
            var sw = new Stopwatch();

            sw.Start();

            using (var connection = new ExcelConnection(string.Format(ConnectionStringTemplate, LargeExcelFile)))
            {
                connection.Open();

                using (var cmd = connection.CreateCommand())
                {
                    cmd.CommandText = "SELECT 1";
                    var value = cmd.ExecuteScalar();

                    Assert.AreEqual(1, value);
                }

                connection.Close();
            }

            sw.Stop();

            Console.WriteLine("Test finished in {0} ms.", sw.ElapsedMilliseconds);
        }
    }
}
