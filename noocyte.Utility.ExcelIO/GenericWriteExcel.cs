using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeOpenXml;
using OfficeOpenXml.Table;

namespace noocyte.Utility.ExcelIO
{
    public class GenericWriteExcel
    {
        public MemoryStream CreateExcelFile<T>(IEnumerable<T> rows) where T : class
        {
            var stream = new MemoryStream();
            var outputStream = new MemoryStream();
            var enumerable = rows as IList<T> ?? rows.ToList();

            using (var package = new ExcelPackage(stream))
            {
                var worksheet = package.Workbook.Worksheets.Add("Sheet1");

                WriteHeader<T>(worksheet);
                WriteContent(worksheet, enumerable);
                FormatSheet(worksheet);

                package.SaveAs(outputStream);
            }
            return outputStream;
        }

        private static void FormatSheet(ExcelWorksheet worksheet)
        {
            // calculate range with data
            var headerAddress = String.Format("{0}:{1}", worksheet.Dimension.Start.Address,
                worksheet.Dimension.End.Address);

            // mark data as a table and format as Light 9 style (config?)
            var tbl = worksheet.Tables.Add(new ExcelAddress(headerAddress), "BaseData");
            tbl.TableStyle = EnumUtils.Parse<TableStyles>("Light9");

            // set all columns to be automatically correct width (perhaps not good idea?)
            worksheet.Cells[headerAddress].AutoFitColumns();
        }

        private static void WriteHeader<T>(ExcelWorksheet worksheet) where T : class
        {
            var beginRow = worksheet.Dimension == null ? 1 : worksheet.Dimension.End.Row;

            var i = 0;
            foreach (var header in HeaderGenerator<T>.GenerateHeaders())
            {
                i++;
                worksheet.Cells[beginRow, i].Value = header.Text;
            }
        }

        private static void WriteContent<T>(ExcelWorksheet worksheet, IEnumerable<T> rows) where T : class
        {
            var rowCount = worksheet.Dimension.End.Row;
            foreach (var rule in rows)
            {
                rowCount++;
                var columnCount = 0;
                foreach (var columnValue in rule.GetPropertyValues())
                {
                    columnCount++;
                    worksheet.Cells[rowCount, columnCount].Value = columnValue.PropertyValue;
                }
            }
        }
    }
}