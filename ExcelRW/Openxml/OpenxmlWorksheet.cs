using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using ExcelReadAndWrite.StdExcelModel;

namespace ExcelReadAndWrite.Openxml
{
    public class OpenxmlWorksheet : StdExcelWorkSheetBase
    {
        public override StdExcelCellBase GetCell(int rowNum, int columnNum)
        {
            throw new NotImplementedException();
        }

        public override string GetCellFormular(int rowNum, int columnNum)
        {
            throw new NotImplementedException();
        }

        public override string GetCellValue(int rowNumber, int columNumber)
        {
            throw new NotImplementedException();
        }

        public override StdExcelColumnBase GetColumn(int index)
        {
            throw new NotImplementedException();
        }

        public override StdExcelRangeBase GetRange()
        {
            throw new NotImplementedException();
        }

        public override StdExcelRowBase GetRow(int index)
        {
            throw new NotImplementedException();
        }

        public override DataTable GetTableContent()
        {
            throw new NotImplementedException();
        }

        public override void InsertColumn(int index)
        {
            throw new NotImplementedException();
        }

        public override void InsertRow(int index)
        {
            throw new NotImplementedException();
        }

        public override void MergeCell(StdExcelRangeBase range)
        {
            throw new NotImplementedException();
        }

        public override void MergeCell(int startRow, int startCol, int endRow, int endCol)
        {
            throw new NotImplementedException();
        }

        public override void SetCellColor(int rowNum, int columnNum, Color color)
        {
            throw new NotImplementedException();
        }

        public override void SetCellFormular(string formular, int rowNum, int columnNum)
        {
            throw new NotImplementedException();
        }

        public override void SetCellValue(string value, int rowNum, int columnNum)
        {
            throw new NotImplementedException();
        }

        public override void SetRangeColor(StdExcelRangeBase range, Color color)
        {
            throw new NotImplementedException();
        }
    }
}
