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
        private Application _ExcelApp;
        private IntPtr _excelProcessHand;
        private string _fileName;
        #endregion

        #region Constructor
        public ComWorkbook():base()
        {
            _ExcelApp = new Application();
            IntPtr _excelProcessHand = new IntPtr(_ExcelApp.Hwnd);
            
            _workbook = _ExcelApp.Workbooks.Add();
        }

        public ComWorkbook(string fileName):base(fileName)
        {
            

        }


        #endregion

        #region Dispose

        ~ComWorkbook()
        {
            int k = 0;
            GetWindowThreadProcessId(_excelProcessHand, out k);
            if (k != 0)
            {
                System.Diagnostics.Process p = System.Diagnostics.Process.GetProcessById(k);

                p.Kill();
            }


        }


        [DllImport("User32.dll", CharSet = CharSet.Auto)]
        private static extern int GetWindowThreadProcessId(IntPtr hwnd, out int ID);

        private void DisposeObj(object o)
        {
            try
            {
                Marshal.ReleaseComObject(o);
            }
            catch (Exception)
            {

                //throw;
            }
            finally
            {
                o = null;
            }
        }

        protected override void Dispose(bool disposing)
        {
            if (disposed)
                return;
            if (_workbook != null)
                _workbook.Close(true);
            if (_ExcelApp != null)
                _ExcelApp.Quit();
            DisposeObj(_workbook);
            DisposeObj(_ExcelApp.Workbooks);
            DisposeObj(_ExcelApp);
            GC.Collect();
           
        }
       
        #endregion

        #region Protected Functions
        protected override void PrePare()
        {
            _ExcelApp = new Application();
            _excelProcessHand = new IntPtr(_ExcelApp.Hwnd);
            
        }

        protected override void ReadWorkbook(string fileName)
        {

            _fileName = fileName;
            _workbook = _ExcelApp.Workbooks.Open(fileName);
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
            _ExcelApp.DisplayAlerts = false;
            _ExcelApp.AlertBeforeOverwriting = false;
            _ExcelApp.Visible = false;
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
