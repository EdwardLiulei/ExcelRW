using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using ExcelReadAndWrite.StdExcelModel;
using System.Data;


namespace ExcelReadAndWrite.Com
{
    public class ComWorksheet:StdExcelWorkSheetBase
    {
        #region Field
        private Worksheet _worksheet;

        #endregion

        #region Constructor
        public ComWorksheet(Worksheet worksheet)
        {
            _worksheet = worksheet;

            _sheetName = worksheet.Name;
        }

        #endregion

        public Worksheet GetComWorksheet()
        {
            return _worksheet;
        }
  
        public override System.Data.DataTable GetTableContent()
        {
            System.Data.DataTable table = new System.Data.DataTable();
            int iRowCount = _worksheet.UsedRange.Rows.Count;
            int iColCount = _worksheet.UsedRange.Columns.Count;
            
            var a = _worksheet.Range[_worksheet.Cells[1,1],_worksheet.Cells[iRowCount,iColCount]].Value2;

            return table;
        }

        public override string GetCellValue(int rowNumber, int columNumber)
        {
            return _worksheet.Cells[rowNumber, columNumber];
        }


        public override StdExcelRangeBase GetRange()
        {
            return null;
        }

        public override StdExcelCellBase GetCell(int rowNum, int columnNum)
        {
            Range cell = _worksheet.Cells[rowNum, columnNum];
            return new ComExcelCell(cell);
        }

        public override string GetCellFormular(int rowNum, int columnNum)
        {
            return null;
        }

        public override StdExcelRowBase GetRow(int index)
        {
            return null;
        }

        public override StdExcelColumnBase GetColumn(int index)
        {
            return null;
        }

        public override void InsertRow(int index)
        { }

        public override void InsertColumn(int index)
        { }

        public override void SetCellValue(string value, int rowNum, int columnNum)
        { }

        public override void SetCellFormular(string formular, int rowNum, int columnNum)
        { }

        public override void SetRangeColor(StdExcelRangeBase range, System.Drawing.Color color)
        { }

        public override void SetCellColor(int rowNum, int columnNum, System.Drawing.Color color)
        { }

        public override void MergeCell(StdExcelRangeBase range) { }

        public override void MergeCell(int startRow, int startCol, int endRow, int endCol) { }
    }
}
