using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelReadAndWrite.StdExcelModel;
using Microsoft.Office.Interop.Excel;
using System.Linq;

namespace ExcelReadAndWrite.Com
{
    public class ComWorkbook:StdExcelWorkbookBase
    {
        #region Field
        private Workbook _workbook;
        private Application _xApp;
        #endregion

        public override StdExcelWorkSheetBase GetSheet(string sheetName)
        {
            StdExcelWorkSheetBase targetSheet =null;
            foreach (Worksheet worksheet in _workbook.Worksheets)
            {
                if (worksheet.Name.Equals(sheetName, StringComparison.OrdinalIgnoreCase))
                    targetSheet = new ComWorksheet(worksheet);
            }
            return targetSheet;
        }

        public override void Load(string fileName)
        {
            _xApp = new Application();
            _workbook = _xApp.Workbooks.Open(fileName);
            foreach (Worksheet sheet in _workbook.Worksheets)
            {
                ComWorksheet comWorksheet = new ComWorksheet(sheet);
                _workSheets.Add(comWorksheet);
            }
            //throw new NotImplementedException();
        }

        public override void Save(string fileName)
        {
            _workbook.SaveAs(fileName);

            //ReleaseReSource();
            //throw new NotImplementedException();
        }

        ~ComWorkbook()
        {
            if (_workbook == null)
                return;
            foreach (var worksheet in _workSheets)
            {
                Worksheet xSheet = ((ComWorksheet)worksheet).GetComWorksheet();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xSheet);
                xSheet = null;
            }

            System.Runtime.InteropServices.Marshal.ReleaseComObject(_workbook);
            _workbook = null;
            _xApp.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(_xApp);

            _xApp = null;
            GC.Collect();
        }

       
    }
}
