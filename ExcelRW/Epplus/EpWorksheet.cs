using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelReadAndWrite.Base;
using OfficeOpenXml;

namespace ExcelReadAndWrite.Epplus
{
    public class EpWorksheet:ExcelWorkSheetBase
    {
        private ExcelWorksheet _worksheet;

        public EpWorksheet(ExcelWorksheet worksheet)
        {
            _worksheet = worksheet;
        }

        public ExcelWorksheet GetEpWorksheet()
        {
            return _worksheet;
        }
    }
}
