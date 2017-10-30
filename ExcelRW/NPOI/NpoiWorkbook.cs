using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelReadAndWrite.StdExcelModel;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using System.IO;
using ExcelReadAndWrite.Util;

namespace ExcelReadAndWrite.NPOI
{
    public class NpoiWorkbook:StdExcelWorkbookBase
    {
        #region Field
        private IWorkbook _npoiWorkbook;

        #endregion

        #region Constructor

        public NpoiWorkbook(WorkBookType type) : base()
        {
            if (type == WorkBookType.XLS)
                _npoiWorkbook = new HSSFWorkbook();
            else
                _npoiWorkbook = new XSSFWorkbook();
        }

        public NpoiWorkbook(string fileName) : base(fileName)
        {

        }

        #endregion

        #region Protected Funtions
        protected override void ReadWorkbook(string fileName)
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

        #endregion

        #region public functions

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

        public override StdExcelWorkSheetBase InsertSheet(string sheetName)
        {
            ISheet sheet = _npoiWorkbook.CreateSheet(sheetName);
            NpoiWorksheet npoiSheet =  new NpoiWorksheet(sheet);
            _workSheets.Add(npoiSheet);
            return npoiSheet;
        }

        #endregion

        public override StdExcelWorkSheetBase InsertSheet(System.Data.DataTable table, string sheetName, bool withHeader)
        {
            throw new NotImplementedException();
        }
    }
}
