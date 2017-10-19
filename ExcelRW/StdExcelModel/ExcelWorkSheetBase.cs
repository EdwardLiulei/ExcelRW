using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using ExcelReadAndWrite.StdExcelModel.BaseModel;

namespace ExcelReadAndWrite.StdExcelModel
{
    public abstract class ExcelWorkSheetBase:IStdWorksheet
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


        public string GetSheetName()
        {
            throw new NotImplementedException();
        }

        public IStdRange GetRange()
        {
            throw new NotImplementedException();
        }

        public IStdCell GetCell(int rowNum, int columnNum)
        {
            throw new NotImplementedException();
        }

        public string GetCellFormular(int rowNum, int columnNum)
        {
            throw new NotImplementedException();
        }

        public IStdRow GetRow(int index)
        {
            throw new NotImplementedException();
        }

        public IStdColumn GetColumn(int index)
        {
            throw new NotImplementedException();
        }

        public void InsertRow(int index)
        {
            throw new NotImplementedException();
        }

        public void InsertColumn(int index)
        {
            throw new NotImplementedException();
        }

        public void SetCellValue(string value, int rowNum, int columnNum)
        {
            throw new NotImplementedException();
        }

        public void SetCellFormular(string formular, int rowNum, int columnNum)
        {
            throw new NotImplementedException();
        }

        public void SetRangeColor(IStdRange range, System.Drawing.Color color)
        {
            throw new NotImplementedException();
        }

        public void SetCellColor(int rowNum, int columnNum, System.Drawing.Color color)
        {
            throw new NotImplementedException();
        }

        public void MergeCell(IStdRange range)
        {
            throw new NotImplementedException();
        }

        public void MergeCell(int startRow, int startCol, int endRow, int endCol)
        {
            throw new NotImplementedException();
        }
    }

}
