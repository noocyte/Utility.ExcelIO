using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeOpenXml;

namespace noocyte.Utility.ExcelIO
{
    public class GenericReadExcel
    {
        public bool HasHeader { get; set; }

        public bool ExtractHeader { get; set; }

        public IDictionary<string, IEnumerable<T>> ExtractRows<T>(Stream stream, Func<Dictionary<int, string>, ExcelWorksheetRow, T> rowFunction)
        {
            using (var package = new ExcelPackage(stream))
            {
                return package.Workbook.Worksheets
                    .ToDictionary(sheet => sheet.Name, sheet => ExtractRows(stream, rowFunction, sheet.Name));
            }
        }

        public IEnumerable<T> ExtractRows<T>(Stream stream, Func<Dictionary<int, string>, ExcelWorksheetRow, T> rowFunction, string sheetName)
        {
            using (var package = new ExcelPackage(stream))
            {
                var worksheet = package.Workbook.Worksheets[sheetName];
                return IterateOverRows(worksheet, rowFunction);
            }
        }

        // ReSharper disable once MethodOverloadWithOptionalParameter
        public IEnumerable<T> ExtractRows<T>(Stream stream, Func<Dictionary<int, string>, ExcelWorksheetRow, T> rowFunction, int sheetNumber = 1)
        {
            using (var package = new ExcelPackage(stream))
            {
                var worksheet = package.Workbook.Worksheets[sheetNumber];
                return IterateOverRows(worksheet, rowFunction);
            }
        }

        private IEnumerable<T> IterateOverRows<T>(ExcelWorksheet worksheet, Func<Dictionary<int, string>, ExcelWorksheetRow, T> rowFunction)
        {
            var beginRow = 1;
            var headerValues = new Dictionary<int, string>();

            if (HasHeader)
            {
                var rowRange = worksheet.Cells[1, 1];
                var row = new ExcelWorksheetRow(rowRange);

                for (var i = 1; i <= worksheet.Dimension.End.Column; i++)
                    headerValues[i] = row.GetValue<string>(i);

                if (!ExtractHeader)
                    beginRow = 2;
            }

            for (var i = beginRow; i <= worksheet.Dimension.End.Row; i++)
            {
                var rowRange = worksheet.Cells[i, 1];
                var row = new ExcelWorksheetRow(rowRange);
                yield return rowFunction(headerValues, row);
            }
        }
    }
}