using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelReadAndWrite.Base
{
    public abstract class ExcelWorkbookBase
    {
        #region Filed
        protected List<ExcelWorkSheetBase> _workSheets;

        #endregion

        #region Properity

        public List<ExcelWorkSheetBase> WorkSheets { get { return _workSheets; } }
        #endregion

        #region Absrtract Fuctions

        public abstract void LoadWorkBook(string fileName);

        public abstract void SaveWorkBook(string fileName);

        public abstract ExcelWorkSheetBase GetSheet(string sheetName);

        public abstract void SaveWorkSheet();
        #endregion
    }
}
