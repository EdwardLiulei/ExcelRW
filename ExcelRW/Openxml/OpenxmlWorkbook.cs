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

        public override void Load(string fileName)
        {
            var xmlDoc = SpreadsheetDocument.Open(fileName, true);
            _xmlWorkbook =  xmlDoc.WorkbookPart;
            var sheets = _xmlWorkbook.Workbook.Descendants<Sheet>();
            foreach (Sheet sheet in _xmlWorkbook.Workbook.Descendants<Sheet>())
            {
                WorksheetPart worksheet = (WorksheetPart)_xmlWorkbook.GetPartById(sheet.Id);

            }
            
        }

        public override void Save(string fileName)
        {
            throw new NotImplementedException();
        }
    }
}
