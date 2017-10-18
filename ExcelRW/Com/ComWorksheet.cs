using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using ExcelReadAndWrite.Base;

namespace ExcelReadAndWrite.Com
{
    public class ComWorksheet:ExcelWorkSheetBase
    {
        public ComWorksheet(Worksheet worksheet)
        {
            _worksheet = worksheet;
 
        }

        private Worksheet _worksheet;

        public Worksheet GetComWorksheet()
        {
            return _worksheet;
        }
    }
}
