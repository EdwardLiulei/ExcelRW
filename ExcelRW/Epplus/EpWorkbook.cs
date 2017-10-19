using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelReadAndWrite.StdExcelModel;
using OfficeOpenXml;
using System.IO;


namespace ExcelReadAndWrite.Epplus
{
    public class EpWorkbook:ExcelWorkbookBase
    {

        #region
        private ExcelWorkbook _workbook;

        public override ExcelWorkSheetBase GetSheet(string sheetName)
        {
            throw new NotImplementedException();
        }

        #endregion

        public override void Load(string fileName)
        {
            if (!File.Exists(fileName))
                throw new Exception(string.Format("The file: {0} does not exists",fileName));
            ExcelPackage package = new ExcelPackage(new FileInfo( fileName));
            _workbook = package.Workbook;
            foreach (var worksheet in _workbook.Worksheets)
            {
                EpWorksheet epWorksheet = new EpWorksheet(worksheet);
                _workSheets.Add(epWorksheet);
            }
        }

        public override void Save(string fileName)
        {
            ExcelPackage package = new ExcelPackage();
            
            foreach (var worksheet in _workSheets)
            {
                var epWorksheet = worksheet as EpWorksheet;
                package.Workbook.Worksheets.Add(worksheet.SheetName,epWorksheet.GetEpWorksheet());
            }
            package.SaveAs(new FileInfo( fileName));
            
        }

    }
}
