#region Copyright

/*
	Copyright (c) Sherzod Mutalov, 2015
	mailto:shmutalov@gmail.com
*/

#endregion

using System.Diagnostics;
using NUnit.Framework;

namespace System.Data.Excel.Test
{
    [TestFixture]
    public class ExcelProviderTests
    {
        private const string SmallExcelXmlFile = @"C:\Sample - Superstore.xlsx";
        private const string SmallExcelBinaryFile = @"C:\Sample - Superstore.xls";
        private const string LargeExcelFile = @"C:\AWLArge_FactInternetSales_1M.xlsx";

        private const string ConnectionStringBinaryTemplates = "Database={0};Type=Binary;Password=;FirstRowIsHeader=True;StorageDir=;ForceStorageReload=True;AnalysisMethod=BestMatch;RowsToAnalyse=50";
        private const string ConnectionStringXmlTemplate = "Database={0};Type=Xml;Password=;FirstRowIsHeader=True;StorageDir=;ForceStorageReload=True;AnalysisMethod=BestMatch;RowsToAnalyse=50";

        [Test]
        public void TestExcelConnectionToSmallXmlFile()
        {
            var sw = new Stopwatch();

            sw.Start();

            using (var connection = new ExcelConnection(string.Format(ConnectionStringXmlTemplate, SmallExcelXmlFile)))
            {
                connection.Open();

                using (var cmd = connection.CreateCommand())
                {
                    cmd.CommandText = "SELECT 1 FROM `Orders` LIMIT 1";
                    var value = cmd.ExecuteScalar();

                    Assert.AreEqual(1, value);
                }

                connection.Close();
            }

            sw.Stop();

            Console.WriteLine("Test finished in {0} ms.", sw.ElapsedMilliseconds);
        }

        [Test]
        public void TestExcelConnectionToSmallBinaryFile()
        {
            var sw = new Stopwatch();

            sw.Start();

            using (var connection = new ExcelConnection(string.Format(ConnectionStringBinaryTemplates, SmallExcelBinaryFile)))
            {
                connection.Open();

                using (var cmd = connection.CreateCommand())
                {
                    cmd.CommandText = "SELECT 1 FROM `Orders` LIMIT 1";
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

            using (var connection = new ExcelConnection(string.Format(ConnectionStringXmlTemplate, LargeExcelFile)))
            {
                connection.Open();

                using (var cmd = connection.CreateCommand())
                {
                    cmd.CommandText = "SELECT 1 FROM `AWLArge_FactInternetSales_1M` LIMIT 1";
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
