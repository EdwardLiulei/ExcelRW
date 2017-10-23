using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using ExcelReadAndWrite.StdExcelModel.BaseModel;

namespace ExcelReadAndWrite.StdExcelModel
{
    public abstract class StdExcelWorkSheetBase
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
        public StdExcelWorkSheetBase()
        {

        }
        #endregion

        #region Abstract Functions

        public abstract string GetCellValue(int rowNumber,int columNumber);

        public abstract DataTable GetTableContent();

        public abstract StdExcelRangeBase GetRange(int startRow, int startCol, int endRow, int endCol);

        public abstract StdExcelCellBase GetCell(int rowNum, int columnNum);

        public abstract string GetCellFormula(int rowNum, int columnNum);

        public abstract StdExcelRowBase GetRow(int index);

        public abstract StdExcelColumnBase GetColumn(int index);
      
        public abstract void InsertRow(int index);

        public abstract void InsertColumn(int index);

        public abstract void SetCellValue(string value, int rowNum, int columnNum);

        public abstract void SetCellFormula(string formular, int rowNum, int columnNum);

        public abstract void SetRangeColor(StdExcelRangeBase range, System.Drawing.Color color);

        public abstract void SetCellColor(int rowNum, int columnNum, System.Drawing.Color color);

        public abstract void MergeCell(StdExcelRangeBase range);

        public abstract void MergeCell(int startRow, int startCol, int endRow, int endCol);
      
        #endregion


        public string GetSheetName()
        {
            return _sheetName;
        }

        
    }

}
