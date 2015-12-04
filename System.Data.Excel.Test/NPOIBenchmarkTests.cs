using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPOI.OpenXml4Net.OPC;
using NPOI.SS.UserModel;
using NPOI.XSSF.Model;
using NPOI.XSSF.UserModel;
using NPOI.XSSF.Extractor;
using NUnit.Framework;

namespace System.Data.Excel.Test
{
    [TestFixture]
    public class NpoiBenchmarkTests
    {
        private const string SmallExcelFile = @"C:\Superstore_Locked.xlsx";
        private const string LargeExcelFile = @"C:\AWLArge_FactInternetSales_1M.xlsx";

        [Test]
        public void TestNpoiLoadingSmallFile()
        {
            var sw = new Stopwatch();
            var totalRowsCount = 0;

            sw.Start();

            var wb = WorkbookFactory.Create(SmallExcelFile);

            for (var sheetId = 0; sheetId < wb.NumberOfSheets; sheetId++)
            {
                var rowsCount = 0;
                var sheet = wb.GetSheetAt(sheetId);

                Console.WriteLine("Processing sheet '{0}'", sheet.SheetName);

                foreach (XSSFRow row in sheet)
                {
                    rowsCount++;
                }

                Console.WriteLine("Processed {0} rows", rowsCount);

                totalRowsCount += rowsCount;
            }

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

            var wb = WorkbookFactory.Create(LargeExcelFile);

            for (var sheetId = 0; sheetId < wb.NumberOfSheets; sheetId++)
            {
                var rowsCount = 0;
                var sheet = wb.GetSheetAt(sheetId);
                
                Console.WriteLine("Processing sheet '{0}'", sheet.SheetName);
                
                foreach (XSSFRow row in sheet)
                {
                    rowsCount++;
                }

                Console.WriteLine("Processed {0} rows", rowsCount);

                totalRowsCount += rowsCount;
            }
            
            sw.Stop();

            Console.WriteLine("Test finished in {0} ms", sw.ElapsedMilliseconds);
            Console.WriteLine("Processed {0} rows", totalRowsCount);
        }
    }
}
