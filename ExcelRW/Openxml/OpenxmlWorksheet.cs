using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using ExcelReadAndWrite.StdExcelModel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelReadAndWrite.Openxml
{
    public class OpenxmlWorksheet : StdExcelWorkSheetBase
    {
        #region Field
        private WorksheetPart _openxmlWorksheet;
        #endregion

        #region Constructor
        public OpenxmlWorksheet(WorksheetPart worksheet)
        {
            _openxmlWorksheet = worksheet;
        }
        #endregion

        public override StdExcelCellBase GetCell(int rowNum, int columnNum)
        {
            IEnumerable<Row> rows = _openxmlWorksheet.Worksheet.Descendants<Row>();
            Row row = rows.ToList()[rowNum];
            IEnumerable<Cell> cells = row.Descendants<Cell>();
            Cell cell = cells.ToList()[columnNum];
            return null;
        }
        

        public override string GetCellFormula(int rowNum, int columnNum)
        {
            IEnumerable<Row> rows = _openxmlWorksheet.Worksheet.Descendants<Row>();
            Row row = rows.ToList()[rowNum];
            IEnumerable<Cell> cells = row.Descendants<Cell>();
            Cell cell = cells.ToList()[columnNum];
            return cell.CellFormula.Text;
        }

        public override string GetCellValue(int rowNum, int columnNum)
        {
            IEnumerable<Row> rows = _openxmlWorksheet.Worksheet.Descendants<Row>();
            Row row = rows.ToList()[rowNum];
            IEnumerable<Cell> cells = row.Descendants<Cell>();
            Cell cell = cells.ToList()[columnNum];
            return cell.CellValue.Text;
        }

        public override StdExcelColumnBase GetColumn(int index)
        {
            throw new NotImplementedException();
        }

        public override StdExcelRangeBase GetRange(int startRow, int startCol, int endRow, int endCol)
        {
            throw new NotImplementedException();
        }

        public override StdExcelRowBase GetRow(int index)
        {
            throw new NotImplementedException();
        }

        public override DataTable GetTableContent(bool hasHeader = false)
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

        public override void SetCellColor(int rowNum, int columnNum, System.Drawing.Color color)
        {
            throw new NotImplementedException();
        }

        public override void SetCellFormula(string formular, int rowNum, int columnNum)
        {
            throw new NotImplementedException();
        }

        public override void SetCellValue(string value, int rowNum, int columnNum)
        {
            throw new NotImplementedException();
        }

        public override void SetRangeColor(StdExcelRangeBase range, System.Drawing.Color color)
        {
            throw new NotImplementedException();
        }

    }
}
