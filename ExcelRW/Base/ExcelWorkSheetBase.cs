using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;


namespace ExcelReadAndWrite.Base
{
    public abstract class ExcelWorkSheetBase
    {
        #region Field

        protected string _sheetName;

        protected DataTable _tableContent;

        #endregion

        #region Properity

        public string SheetName 
        {
            get { return _sheetName; }
        }

        #endregion

        #region Constructor
        public ExcelWorkSheetBase()
        {

        }
        #endregion

        #region Abstract Functions

        public abstract string GetCellValue(int rowNumber,int columNumber);

        public abstract DataTable GetTableContent();

        #endregion

    }

}
