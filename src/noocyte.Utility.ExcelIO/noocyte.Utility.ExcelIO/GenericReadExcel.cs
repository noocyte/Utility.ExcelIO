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

        public IDictionary<string, IEnumerable<T>> ExtractRows<T>(Stream stream, Func<ExcelWorksheetRow, T> rowFunction)
        {
            using (var package = new ExcelPackage(stream))
            {
                return package.Workbook.Worksheets
                    .ToDictionary(sheet => sheet.Name, sheet => ExtractRows(stream, rowFunction, sheet.Name));
            }
        }

        public IEnumerable<T> ExtractRows<T>(Stream stream, Func<ExcelWorksheetRow, T> rowFunction, string sheetName)
        {
            using (ExcelPackage package = new ExcelPackage(stream))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[sheetName];
                return IterateOverRows(worksheet, rowFunction);
            }
        }

        public IEnumerable<T> ExtractRows<T>(Stream stream, Func<ExcelWorksheetRow, T> rowFunction, int sheetNumber = 1)
        {
            using (ExcelPackage package = new ExcelPackage(stream))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[sheetNumber];
                return IterateOverRows(worksheet, rowFunction);
            }
        }

        private IEnumerable<T> IterateOverRows<T>(ExcelWorksheet worksheet, Func<ExcelWorksheetRow, T> rowFunction)
        {
            int beginRow = 1;
            if (HasHeader && !ExtractHeader)
                beginRow = 2;

            for (int i = beginRow; i <= worksheet.Dimension.End.Row; i++)
            {
                var rowRange = worksheet.Cells[i, 1];
                var row = new ExcelWorksheetRow(rowRange);
                yield return rowFunction(row);
            }
        }
    }
}