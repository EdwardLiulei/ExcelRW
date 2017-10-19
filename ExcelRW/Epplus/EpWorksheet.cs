using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using ExcelReadAndWrite.StdExcelModel;
using OfficeOpenXml;

namespace ExcelReadAndWrite.Epplus
{
    public class EpWorksheet:ExcelWorkSheetBase
    {
        private ExcelWorksheet _worksheet;

        public EpWorksheet(ExcelWorksheet worksheet)
        {
            _worksheet = worksheet;
            _sheetName = worksheet.Name;
        }

        public override string GetCellValue(int rowNumber, int columNumber)
        {
            throw new NotImplementedException();
        }

        public ExcelWorksheet GetEpWorksheet()
        {
            return _worksheet;
        }

        public override DataTable GetTableContent()
        {
            throw new NotImplementedException();
        }
    }
}
