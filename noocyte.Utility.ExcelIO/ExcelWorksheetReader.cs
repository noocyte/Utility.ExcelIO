using OfficeOpenXml;

namespace noocyte.Utility.ExcelIO
{
    public class ExcelWorksheetRow
    {
        public ExcelWorksheetRow(ExcelRange row)
        {
            CurrentRow = row;
            CurrentRowIndex = CurrentRow.Start.Row;
        }

        private ExcelRange CurrentRow { get; set; }
        private int CurrentRowIndex { get; set; }

        public T GetValue<T>(int columnIndex)
        {
            return CurrentRow[CurrentRowIndex, columnIndex].GetValue<T>();
        }
    }
}