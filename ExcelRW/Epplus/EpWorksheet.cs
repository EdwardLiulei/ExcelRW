using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using ExcelReadAndWrite.StdExcelModel;
using OfficeOpenXml;
using ExcelReadAndWrite.Util;

namespace ExcelReadAndWrite.Epplus
{
    public class EpWorksheet:StdExcelWorkSheetBase
    {
        #region Field
        private ExcelWorksheet _epWorksheet;

        #endregion

        #region Properity

        public override int ColumnNum
        {
            
            get
            {
                if (_epWorksheet.Dimension == null)
                    return 1;
                return _epWorksheet.Dimension.End.Column;
            }
        }

        public override int RowNum
        {

            get
            {
                if (_epWorksheet.Dimension == null)
                    return 1;
                return _epWorksheet.Dimension.End.Row;
            }
        }

        #endregion

        #region Constructor

        public EpWorksheet(ExcelWorksheet worksheet)
        {
            _epWorksheet = worksheet;
            _sheetName = worksheet.Name;
        }

        #endregion

        #region Public Functions

        public ExcelWorksheet GetEpWorksheet()
        {
            
            return _epWorksheet;
        }

        public override DataTable GetTableContent(bool hasHeader = false)
        {
            System.Data.DataTable table = new System.Data.DataTable();
            int iRowCount = _epWorksheet.Dimension.End.Row;
            int iColCount = _epWorksheet.Dimension.End.Column;

            if (hasHeader == true)
            {
                List<string> headers = new List<string>();
                for (int i = 1; i <= iColCount; i++)
                {
                    headers.Add(GetCellValue(1, i));
                }
                if (headers.Count > headers.Distinct().Count())
                    throw new Exception(string.Format("The sheet: {0} contains duplicate headers", _sheetName));
                foreach (var header in headers)
                {
                    table.Columns.Add(header);
                }
            }
            else
            {
                for (int i = 0; i < iColCount; i++)
                {
                    string columnName = WorksheetAddress.GetColumnAddress(i + 1);
                    table.Columns.Add(columnName);
                }
                DataRow row = table.NewRow();
                for (int j = 1; j <= iColCount; j++)
                {
                    row[j-1] = GetCellValue(1, j);
                }
                table.Rows.Add(row);
            }


            for (int i = 2; i <= iRowCount; i++)
            {
                DataRow row = table.NewRow();
                for (int j = 1; j <= iColCount; j++)
                {

                    row[j-1] = GetCellValue(i, j);
                }
                table.Rows.Add(row);
            }
            return table;
        }


        public override string GetCellValue(int rowNumber, int columNumber)
        {
            if (_epWorksheet.Cells[rowNumber, columNumber].Value == null)
                return "";
            else
                return _epWorksheet.Cells[rowNumber, columNumber].Value.ToString();
        }


        public override ExcelReadAndWrite.StdExcelModel.StdExcelRangeBase GetRange(int startRow, int startCol, int endRow, int endCol)
        {
            ExcelRange range = _epWorksheet.SelectedRange[startRow, startCol, endRow, endCol];

            return new EpExcelRange(range);
        }

        public override ExcelReadAndWrite.StdExcelModel.StdExcelCellBase GetCell(int rowNum, int columnNum)
        {
            ExcelRange range = _epWorksheet.Cells[rowNum, columnNum];
            
            return new EpExcelCell(range);
        }

        public override string GetCellFormula(int rowNum, int columnNum)
        {
            if (_epWorksheet.Cells[rowNum, columnNum].Value == null)
                return "";
            else
                return _epWorksheet.Cells[rowNum, columnNum].Formula;
        }

        public override StdExcelRowBase GetRow(int index)
        {
            //var row= _comWorksheet.Row(index);
            return new EpExcelRow(_epWorksheet,index);
        }

        public override StdExcelColumnBase GetColumn(int index)
        {
           

            return new EpExcelColumn(_epWorksheet,index);
        }

        public override void InsertRow(int index)
        {
            _epWorksheet.InsertRow(index,index);
        }

        public override void InsertColumn(int index)
        {
            _epWorksheet.InsertColumn(index, index);
        }

        public override void SetCellValue(string value, int rowNum, int columnNum)
        {
            var cell = GetCell(rowNum, columnNum);
            cell.SetValue(value);
        }

        public override void SetCellFormula(string formular, int rowNum, int columnNum)
        {
            var cell = GetCell(rowNum, columnNum);
            cell.SetFormular(formular);
        }

        public override void SetRangeColor(ExcelReadAndWrite.StdExcelModel.StdExcelRangeBase range, System.Drawing.Color color)
        {
            range.SetBackgroudColor(color);
        }

        public override void SetCellColor(int rowNum, int columnNum, System.Drawing.Color color)
        {
            var cell = GetCell(rowNum, columnNum);
            cell.SetBackgroudColor(color);
        }

        public override void MergeCell(ExcelReadAndWrite.StdExcelModel.StdExcelRangeBase range)
        {
            range.SetMerge();
        }

        public override void MergeCell(int startRow, int startCol, int endRow, int endCol)
        {
            StdExcelRangeBase range = GetRange(startRow, startCol, endRow, endCol);
            range.SetMerge();
        }

        public override List<string> GetSheetDataFromRow(int rowNum)
        {
            List<string> rowData = new List<string>();

            int columnNum = _epWorksheet.Dimension.End.Column;
            for (int i = 1; i <= columnNum; i++)
            {
                rowData.Add(GetCellValue(rowNum, columnNum));
            }

            return rowData;
        }

        public override List<string> GetSheetDataFromColumn(int columnNum)
        {
            List<string> columnData = new List<string>();

            int rowNum = _epWorksheet.Dimension.End.Row;
            for (int i = 1; i <= columnNum; i++)
            {
                columnData.Add(GetCellValue(rowNum, columnNum));
            }

            return columnData;
        }

        #endregion
    }
}
