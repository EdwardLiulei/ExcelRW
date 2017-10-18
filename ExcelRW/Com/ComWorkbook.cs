using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelReadAndWrite.Base;
using Microsoft.Office.Interop.Excel;

namespace ExcelReadAndWrite.Com
{
    public class ComWorkbook:ExcelWorkbookBase
    {
        private Workbook _workbook;


        public override void LoadWorkBook(string fileName)
        {
            Application xApp = new Application();
            Workbook _workbook = xApp.Workbooks.Open(fileName);
            //throw new NotImplementedException();
        }

        public override void SaveWorkBook(string fileName)
        {
            _workbook.SaveAs(fileName);
            //throw new NotImplementedException();
        }
    }
}
