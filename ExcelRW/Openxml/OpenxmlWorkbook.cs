using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelReadAndWrite.StdExcelModel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelReadAndWrite.Openxml
{
    public class OpenxmlWorkbook : StdExcelWorkbookBase
    {
        private WorkbookPart _xmlWorkbook;
        public override StdExcelWorkSheetBase GetSheet(string sheetName)
        {
            throw new NotImplementedException();
        }

        protected override void ReadWorkbook(string fileName)
        {
            var xmlDoc = SpreadsheetDocument.Open(fileName, true);
            _xmlWorkbook =  xmlDoc.WorkbookPart;
            var sheets = _xmlWorkbook.Workbook.Descendants<Sheet>();
            foreach (Sheet sheet in _xmlWorkbook.Workbook.Descendants<Sheet>())
            {
                WorksheetPart worksheet = (WorksheetPart)_xmlWorkbook.GetPartById(sheet.Id);
                StdExcelWorkSheetBase stdWorksheet = new OpenxmlWorksheet(worksheet);
                _workSheets.Add(stdWorksheet);
            }
            
        }

        public override void Save(string fileName)
        {
            throw new NotImplementedException();
        }

        public override StdExcelWorkSheetBase InsertSheet(string sheetName)
        {
            throw new NotImplementedException();
        }

        public override StdExcelWorkSheetBase InsertSheet(System.Data.DataTable table, string sheetName, bool withHeader)
        {
            throw new NotImplementedException();
        }
    }
}
