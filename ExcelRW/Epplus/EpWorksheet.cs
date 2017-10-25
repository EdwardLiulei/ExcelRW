using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using ExcelReadAndWrite.StdExcelModel;
using OfficeOpenXml;

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

        public override DataTable GetTableContent()
        {
            return new DataTable();
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
