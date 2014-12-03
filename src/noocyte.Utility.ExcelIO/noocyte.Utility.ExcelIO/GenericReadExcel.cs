using System;
using System.Collections.Generic;
using System.IO;
using OfficeOpenXml;

namespace noocyte.Utility.ExcelIO
{
    public class GenericReadExcel
    {
        public IEnumerable<T> ExtractRows<T>(Stream stream, Func<ExcelWorksheetRow, T> rowFunction, string sheetName)
        {
            using (ExcelPackage package = new ExcelPackage(stream))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[sheetName];

				foreach (var row in IterateOverRows(worksheet, rowFunction))
				{
					yield return row;
				}
            }
        }

        public IEnumerable<T> ExtractRows<T>(Stream stream, Func<ExcelWorksheetRow, T> rowFunction, int sheetNumber = 1)
        {
            using (ExcelPackage package = new ExcelPackage(stream))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[sheetNumber];

				foreach (var row in IterateOverRows(worksheet, rowFunction))
	            {
		            yield return row;
	            }
            }
        }

        private IEnumerable<T> IterateOverRows<T>(ExcelWorksheet worksheet, Func<ExcelWorksheetRow, T> rowFunction)
        {
            int beginRow = 1;
            if (this.HasHeader && !this.ExtractHeader)
                beginRow = 2;

            for (int i = beginRow; i <= worksheet.Dimension.End.Row; i++)
            {
                var rowRange = worksheet.Cells[i, 1];
                var row = new ExcelWorksheetRow(rowRange);
                yield return rowFunction(row);
            }
        }

        public bool HasHeader { get; set; }

        public bool ExtractHeader { get; set; }
    }
}
