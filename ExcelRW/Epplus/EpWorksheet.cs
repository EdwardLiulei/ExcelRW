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
        private ExcelWorksheet _worksheet;

        public EpWorksheet(ExcelWorksheet worksheet)
        {
            _worksheet = worksheet;
            _sheetName = worksheet.Name;
        }

       

        public ExcelWorksheet GetEpWorksheet()
        {
            return _worksheet;
        }

        public override DataTable GetTableContent(bool hasHeader = false)
        {
            System.Data.DataTable table = new System.Data.DataTable();
            int iRowCount = _worksheet.Dimension.End.Row;
            int iColCount = _worksheet.Dimension.End.Column;

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
            if (_worksheet.Cells[rowNumber, columNumber].Value == null)
                return "";
            else
                return _worksheet.Cells[rowNumber, columNumber].Value.ToString();
        }


        public override ExcelReadAndWrite.StdExcelModel.StdExcelRangeBase GetRange(int startRow, int startCol, int endRow, int endCol)
        {
            ExcelRange range = _worksheet.SelectedRange[startRow, startCol, endRow, endCol];

            return new EpExcelRange(range);
        }

        public override ExcelReadAndWrite.StdExcelModel.StdExcelCellBase GetCell(int rowNum, int columnNum)
        {
            ExcelRange range = _worksheet.Cells[rowNum, columnNum];
            
            return new EpExcelCell(range);
        }

        public override string GetCellFormula(int rowNum, int columnNum)
        {
            if (_worksheet.Cells[rowNum, columnNum].Value == null)
                return "";
            else
                return _worksheet.Cells[rowNum, columnNum].Formula;
        }

        public override StdExcelRowBase GetRow(int index)
        {
            //var row= _worksheet.Row(index);
            return new EpExcelRow(_worksheet,index);
        }

        public override StdExcelColumnBase GetColumn(int index)
        {
           

            return new EpExcelColumn(_worksheet,index);
        }

        public override void InsertRow(int index)
        {
            _worksheet.InsertRow(index,index);
        }

        public override void InsertColumn(int index)
        {
            _worksheet.InsertColumn(index, index);
        }

        public override void SetCellValue(string value, int rowNum, int columnNum)
        { }

        public override void SetCellFormula(string formular, int rowNum, int columnNum)
        { }

        public override void SetRangeColor(ExcelReadAndWrite.StdExcelModel.StdExcelRangeBase range, System.Drawing.Color color)
        { }

        public override void SetCellColor(int rowNum, int columnNum, System.Drawing.Color color)
        { }

        public override void MergeCell(ExcelReadAndWrite.StdExcelModel.StdExcelRangeBase range) { }

        public override void MergeCell(int startRow, int startCol, int endRow, int endCol) { }
    }
}
