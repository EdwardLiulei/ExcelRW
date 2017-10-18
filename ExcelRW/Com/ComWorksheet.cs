using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using ExcelReadAndWrite.Base;
using System.Data;


namespace ExcelReadAndWrite.Com
{
    public class ComWorksheet:ExcelWorkSheetBase
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

        public override string GetCellValue(int rowNumber, int columNumber)
        {
            throw new NotImplementedException();
        }

        public override System.Data.DataTable GetTableContent()
        {
            System.Data.DataTable table = new System.Data.DataTable();
            int iRowCount = _worksheet.UsedRange.Rows.Count;
            int iColCount = _worksheet.UsedRange.Columns.Count;
            
            var a = _worksheet.Range[_worksheet.Cells[1,1],_worksheet.Cells[iRowCount,iColCount]].Value2;

            return table;
        }
    }
}
