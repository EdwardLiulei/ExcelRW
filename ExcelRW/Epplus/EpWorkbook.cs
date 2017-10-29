using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelReadAndWrite.StdExcelModel;
using OfficeOpenXml;
using System.IO;


namespace ExcelReadAndWrite.Epplus
{
    public class EpWorkbook:StdExcelWorkbookBase
    {

        #region Field
        private ExcelWorkbook _epWorkbook;

        public override StdExcelWorkSheetBase GetSheet(string sheetName)
        {
            throw new NotImplementedException();
        }

        #endregion

        #region Constructor

        public EpWorkbook() : base()
        {
            ExcelPackage package = new ExcelPackage();
            _epWorkbook = package.Workbook;
        }

        public EpWorkbook(string fileName) : base(fileName)
        {

        }

        #endregion

        #region Protected Functions
        protected override void ReadWorkbook(string fileName)
        {
            if (!File.Exists(fileName))
                throw new Exception(string.Format("The file: {0} does not exists",fileName));
            ExcelPackage package = new ExcelPackage(new FileInfo( fileName));
            _epWorkbook = package.Workbook;
            foreach (var worksheet in _epWorkbook.Worksheets)
            {
                EpWorksheet epWorksheet = new EpWorksheet(worksheet);
                _workSheets.Add(epWorksheet);
            }
        }

        #endregion

        #region Public Functions

        public override void Save(string fileName)
        {
            ExcelPackage package = new ExcelPackage();
            
            foreach (var worksheet in _workSheets)
            {
                var epWorksheet = worksheet as EpWorksheet;
                package.Workbook.Worksheets.Add(worksheet.SheetName,epWorksheet.GetEpWorksheet());
            }
            package.SaveAs(new FileInfo(fileName));
            
        }

        public override StdExcelWorkSheetBase InsertSheet(string sheetName)
        {
            ExcelWorksheet worksheet = _epWorkbook.Worksheets.Add(sheetName);
            EpWorksheet epworksheet = new EpWorksheet(worksheet);
            _workSheets.Add(epworksheet);
            return epworksheet;
        }

        #endregion
    }
}
