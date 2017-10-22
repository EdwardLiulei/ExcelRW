using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelReadAndWrite.StdExcelModel;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using System.IO;

namespace ExcelReadAndWrite.NPOI
{
    public class NpoiWorkbook:StdExcelWorkbookBase
    {
        private IWorkbook _npoiWorkbook;
        public override void Load(string fileName)
        {
            string extension = Path.GetExtension(fileName);
            FileStream fs = File.OpenRead(fileName);
            if (extension.Equals(".xls"))
            {
                //把xls文件中的数据写入wk中
                _npoiWorkbook = new HSSFWorkbook(fs);
            }
            else
            {
                //把xlsx文件中的数据写入wk中
                _npoiWorkbook = new XSSFWorkbook(fs);
            }
            fs.Close();
            //读取当前表数据
            int num = _npoiWorkbook.NumberOfSheets;
            for (int i = 0; i < num; i++)
            {
                ISheet sheet = _npoiWorkbook.GetSheetAt(i);
                NpoiWorksheet worksheet = new NpoiWorksheet(sheet);
                _workSheets.Add(worksheet);
            }
            //throw new NotImplementedException();
        }

        public override void Save(string fileName)
        {
            using (var wook = new FileStream(fileName, FileMode.Create, FileAccess.Write))
            {
                _npoiWorkbook.Write(wook);
            }
        }

        public override StdExcelWorkSheetBase GetSheet(string sheetName)
        {
            return _workSheets.Find(p => p.SheetName.Equals(sheetName, StringComparison.OrdinalIgnoreCase));
        }
    }
}
