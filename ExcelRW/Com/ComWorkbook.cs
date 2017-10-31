using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelReadAndWrite.StdExcelModel;
using Microsoft.Office.Interop.Excel;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using ExcelReadAndWrite.StdExcelModel.DataModel;

namespace ExcelReadAndWrite.Com
{
    public class ComWorkbook:StdExcelWorkbookBase
    {
        #region Field
        private Workbook _workbook;
        private Application _xApp;
        private string _fileName;
        #endregion

        #region Constructor
        public ComWorkbook():base()
        {
            _xApp = new Application();
            _workbook = _xApp.Workbooks.Add();
        }

        public ComWorkbook(string fileName):base(fileName)
        {
            

        }


        #endregion

        #region Dispose

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

            if (_xApp == null)
                return;
            IntPtr t = new IntPtr(_xApp.Hwnd);
            int k = 0;
            GetWindowThreadProcessId(t, out k);
            System.Diagnostics.Process p = System.Diagnostics.Process.GetProcessById(k);
            p.Kill();

        }


        [DllImport("User32.dll", CharSet = CharSet.Auto)]
        private static extern int GetWindowThreadProcessId(IntPtr hwnd, out int ID);

        #endregion

        #region Protected Functions
        protected override void PrePare()
        {
            _xApp = new Application();
        }

        protected override void ReadWorkbook(string fileName)
        {

            _fileName = fileName;
            _workbook = _xApp.Workbooks.Open(fileName);
            foreach (Worksheet sheet in _workbook.Worksheets)
            {
                ComWorksheet comWorksheet = new ComWorksheet(sheet);
                _workSheets.Add(comWorksheet);
            }
            //throw new NotImplementedException();
        }
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

        #endregion

        #region Public Functions
        public override void Save(string fileName)
        {
            _xApp.DisplayAlerts = false;
            _xApp.AlertBeforeOverwriting = false;
            _xApp.Visible = false;
            if (fileName == _fileName)
                _workbook.Save();
            else
                _workbook.SaveAs(fileName, Type.Missing, Type.Missing,Type.Missing,Type.Missing);

            //ReleaseReSource();
            //throw new NotImplementedException();
        }

        public override StdExcelWorkSheetBase InsertSheet(string sheetName)
        {
            Worksheet sheet = _workbook.Sheets.Add();
            sheet.Name = sheetName;
            StdExcelWorkSheetBase worksheet =  new ComWorksheet(sheet);
            _workSheets.Add(worksheet);
            return worksheet;
        }

        #endregion



        public override StdExcelWorkSheetBase InsertSheet(System.Data.DataTable table, string sheetName, bool withHeader)
        {
            throw new NotImplementedException();
        }
    }
}
