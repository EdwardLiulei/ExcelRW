using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelReadAndWrite.StdExcelModel;
using DataTable = System.Data.DataTable;
using System.Data.OleDb;
using System.Data;

namespace ExcelReadAndWrite.Oledb
{
    public class OledbWorksheet:StdExcelWorkSheetBase
    {
        #region Fied
        private string _connectStr;
        private DataTable _tableContent;

        #endregion


        #region Properity
        public override int ColumnNum => throw new NotImplementedException();

        public override int RowNum => throw new NotImplementedException();

        #endregion
        public OledbWorksheet(DataTable table, string sheetName,string connectStr)
        {
            _sheetName =sheetName;
            _connectStr = connectStr;
            _tableContent = table;
        }

        public override string GetCellValue(int rowNumber, int columNumber)
        {
            throw new NotImplementedException();
        }

        public override System.Data.DataTable GetTableContent(bool hasHeader = false)
        {

            return _tableContent;
        }

        public override StdExcelRangeBase GetRange(int startRow, int startCol, int endRow, int endCol)
        {
            return null;
        }

        public override StdExcelCellBase GetCell(int rowNum, int columnNum)
        {
            return null;
        }

        public override string GetCellFormula(int rowNum, int columnNum)
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

        public override void SetCellFormula(string formular, int rowNum, int columnNum)
        { }

        public override void SetRangeColor(StdExcelRangeBase range, System.Drawing.Color color)
        { }

        public override void SetCellColor(int rowNum, int columnNum, System.Drawing.Color color)
        { }

        public override void MergeCell(StdExcelRangeBase range) { }

        public override void MergeCell(int startRow, int startCol, int endRow, int endCol) { }

        public override List<string> GetSheetDataFromRow(int rowNum)
        {
            throw new NotImplementedException();
        }

        public override List<string> GetSheetDataFromColumn(int columnNum)
        {
            throw new NotImplementedException();
        }
    }
}
