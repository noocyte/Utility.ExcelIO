using OfficeOpenXml;

namespace noocyte.Utility.ExcelIO
{
    public class ExcelWorksheetRow
    {
        private ExcelRange CurrentRow { get; set; }
        private int CurrentRowIndex { get; set; }
        public ExcelWorksheetRow(ExcelRange row)
        {
            this.CurrentRow = row;
            this.CurrentRowIndex = this.CurrentRow.Start.Row;
        }

        public T GetValue<T>(int columnIndex)
        {
            return this.CurrentRow[CurrentRowIndex, columnIndex].GetValue<T>();
        }
    }
}
