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
            MemoryStream stream = null;
            stream = new MemoryStream();
            using (ExcelPackage package = new ExcelPackage(stream))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Sheet1");
                WriteHeader(worksheet, rows.First());
                WriteContent(worksheet, rows);
                FormatSheet(worksheet);

                package.Save();
            }
            return stream;
        }

        private void FormatSheet(ExcelWorksheet worksheet)
        {
            // calculate range with data
            string headerAddress = String.Format("{0}:{1}", worksheet.Dimension.Start.Address, worksheet.Dimension.End.Address);

            // mark data as a table and format as Light 9 style (config?)
            var tbl = worksheet.Tables.Add(new ExcelAddress(headerAddress), "BaseData");
            tbl.TableStyle = EnumUtils.Parse<TableStyles>("Light9");

            // set all columns to be automatically correct width (perhaps not good idea?)
            worksheet.Cells[headerAddress].AutoFitColumns();
        }

        private void WriteHeader<T>(ExcelWorksheet worksheet, T oneRow) where T : class
        {
            int beginRow = worksheet.Dimension == null ? 1 : worksheet.Dimension.End.Row;

            int i = 0;
            foreach (var header in HeaderGenerator<T>.GenerateHeaders())
            {
                i++;
                worksheet.Cells[beginRow, i].Value = header.Text;
            }
        }

        private void WriteContent<T>(ExcelWorksheet worksheet, IEnumerable<T> rows) where T : class
        {
            int columnCount = 0;
            int rowCount = worksheet.Dimension.End.Row;
            foreach (var rule in rows)
            {
                rowCount++;
                columnCount = 0;
                foreach (var columnValue in rule.GetPropertyValues())
                {
                    columnCount++;
                    worksheet.Cells[rowCount, columnCount].Value = columnValue.PropertyValue;
                }
            }
        }
    }
}
