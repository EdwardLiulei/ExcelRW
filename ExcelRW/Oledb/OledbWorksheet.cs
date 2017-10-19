using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelReadAndWrite.StdExcelModel;

namespace ExcelReadAndWrite.Oledb
{
    public class OledbWorksheet:ExcelWorkSheetBase
    {
        private string _connectStr;

        public OledbWorksheet(string sheetName,string connectStr)
        {
            _sheetName =sheetName;
            _connectStr = connectStr;
        }

        public override string GetCellValue(int rowNumber, int columNumber)
        {
            throw new NotImplementedException();
        }

        public override System.Data.DataTable GetTableContent()
        {
            throw new NotImplementedException();
        }
    }
}
