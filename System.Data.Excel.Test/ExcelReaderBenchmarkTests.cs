using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel;
using NUnit.Framework;

namespace System.Data.Excel.Test
{
    [TestFixture]
    public class ExcelReaderBenchmarkTests
    {
        private const string SmallExcelFile = @"C:\Superstore_Locked.xlsx";
        private const string LargeExcelFile = @"C:\AWLArge_FactInternetSales_1M.xlsx";

        [Test]
        public void TestNpoiLoadingSmallFile()
        {
            var sw = new Stopwatch();
            var totalRowsCount = 0;

            sw.Start();

            var reader = ExcelReaderFactory.CreateOpenXmlReader(File.OpenRead(SmallExcelFile));

            do
            {
                var rowsCount = 0;
                Console.WriteLine("Processing sheet '{0}'", reader.Name);

                while (reader.Read())
                {
                    rowsCount++;
                }

                Console.WriteLine("Processed {0} rows", rowsCount);

                totalRowsCount += rowsCount;
            } while (reader.NextResult());

            sw.Stop();

            Console.WriteLine("Test finished in {0} ms", sw.ElapsedMilliseconds);
            Console.WriteLine("Processed {0} rows", totalRowsCount);
        }

        [Test]
        public void TestNpoiLoadingLargeFile()
        {
            var sw = new Stopwatch();
            var totalRowsCount = 0;

            sw.Start();

            var reader = ExcelReaderFactory.CreateOpenXmlReader(File.OpenRead(LargeExcelFile));

            do
            {
                var rowsCount = 0;
                Console.WriteLine("Processing sheet '{0}'", reader.Name);

                while (reader.Read())
                {
                    rowsCount++;
                }

                Console.WriteLine("Processed {0} rows", rowsCount);

                totalRowsCount += rowsCount;
            } while (reader.NextResult());

            sw.Stop();

            Console.WriteLine("Test finished in {0} ms", sw.ElapsedMilliseconds);
            Console.WriteLine("Processed {0} rows", totalRowsCount);
        }
    }
}
