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

        private string _sheetName;

        private DataTable _tableContent;

        private bool _hasHeader;

        #endregion

        #region Properity

        public string SheetName 
        {
            get { return _sheetName; }
        }
        
        #endregion

        #region Abstract Functions

        public abstract string GetCellValue(int rowNumber,int columNumber);

        #endregion

    }

}
