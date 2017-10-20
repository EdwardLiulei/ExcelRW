using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelReadAndWrite.StdExcelModel;
using System.Data.OleDb;

namespace ExcelReadAndWrite.Oledb
{
    public class OledbWorkbook:StdExcelWorkbookBase
    {
        private string _connectStr;

        public override void Load(string fileName)
        {
            string provider = "Provider=Microsoft.ACE.OLEDB.12.0;";
            string dataSource = "Data Source="+fileName+";";
            string extended = "Extended Properties=Excel 8.0;";

            _connectStr = provider + dataSource + extended;
            OleDbConnection objConn = new OleDbConnection(_connectStr);
            objConn.Open();
            System.Data.DataTable dt = objConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

            foreach (System.Data.DataRow row in dt.Rows)
            {
                string sheetName = row["TABLE_NAME"].ToString(); //就是

                _workSheets.Add(new OledbWorksheet(sheetName, _connectStr));
            }
            //throw new NotImplementedException();
        }

        public override void Save(string fileName)
        {
            throw new NotImplementedException();
        }

        public override StdExcelWorkSheetBase GetSheet(string sheetName)
        {
            throw new NotImplementedException();
        }
    }
}
