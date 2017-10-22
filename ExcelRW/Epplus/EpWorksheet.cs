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
            throw new NotImplementedException();
        }


        public override string GetCellValue(int rowNumber, int columNumber)
        {
            return _worksheet.Cells[rowNumber, columNumber].Value.ToString();
        }


        public override ExcelReadAndWrite.StdExcelModel.StdExcelRangeBase GetRange()
        {
            return null;
        }

        public override ExcelReadAndWrite.StdExcelModel.StdExcelCellBase GetCell(int rowNum, int columnNum)
        {
            return null;
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

        public override void SetRangeColor(ExcelReadAndWrite.StdExcelModel.StdExcelRangeBase range, System.Drawing.Color color)
        { }

        public override void SetCellColor(int rowNum, int columnNum, System.Drawing.Color color)
        { }

        public override void MergeCell(ExcelReadAndWrite.StdExcelModel.StdExcelRangeBase range) { }

        public override void MergeCell(int startRow, int startCol, int endRow, int endCol) { }
    }
}
