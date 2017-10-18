using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelReadAndWrite.Base;
using OfficeOpenXml;
using System.IO;


namespace ExcelReadAndWrite.Epplus
{
    public class EpWorkbook:ExcelWorkbookBase
    {

        #region
        private ExcelWorkbook _workbook;
        
        #endregion

        public override void LoadWorkBook(string fileName)
        {
            if (!File.Exists(fileName))
                throw new Exception(string.Format("The file: {0} does not exists",fileName));
            ExcelPackage package = new ExcelPackage(new FileInfo( fileName));
            _workbook = package.Workbook;
        }

        public override void SaveWorkBook(string fileName)
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
