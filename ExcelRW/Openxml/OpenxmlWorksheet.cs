using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using ExcelReadAndWrite.StdExcelModel;

namespace ExcelReadAndWrite.Openxml
{
    public class OpenxmlWorksheet : ExcelWorkSheetBase
    {


        
        public override string GetCellValue(int rowNumber, int columNumber)
        {
            throw new NotImplementedException();
        }

        public override DataTable GetTableContent()
        {
            throw new NotImplementedException();
        }
    }
}
