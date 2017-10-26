using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelReadAndWrite.StdExcelModel.BaseModel;
using System.IO;

namespace ExcelReadAndWrite.StdExcelModel
{
    public abstract class StdExcelWorkbookBase:IStdWorkbook
    {
        #region Filed
        protected List<StdExcelWorkSheetBase> _workSheets;

        protected WorkBookType _type;

        #endregion

        #region Properity

        public List<StdExcelWorkSheetBase> WorkSheets { get { return _workSheets; } }
        #endregion

        #region Constructor
        public StdExcelWorkbookBase()
        {
            _workSheets = new List<StdExcelWorkSheetBase>();
        }
        #endregion

        #region Absrtract Fuctions

        public abstract void Load(string fileName);

        public abstract void Save(string fileName);
        

        public abstract StdExcelWorkSheetBase GetSheet(string sheetName);
        #endregion

        #region Public Functions


        public List<string> GetSheetList()
        {
            return _workSheets.Select(p => p.GetSheetName()).ToList();
        }

        public int GetSheetCount()
        {
            return _workSheets.Count();
        }

        public StdExcelWorkSheetBase GetSheetByName(string sheetName)
        {
            return _workSheets.Find(p => p.GetSheetName().Equals(sheetName,StringComparison.OrdinalIgnoreCase));
        }

        public StdExcelWorkSheetBase GetSheetByIndex(int index)
        {
            return _workSheets[index];
        }

        public string GetSheetNameByIndex(int index)
        {
            return _workSheets[index].GetSheetName();
        }

        public StdExcelWorkSheetBase CloneSheet(int index)
        {
            throw new NotImplementedException();
        }

        public StdExcelWorkSheetBase CloneSheet(string sheetName)
        {
            throw new NotImplementedException();
        }

        public bool Is1904()
        {
            throw new NotImplementedException();
        }

        #endregion

        #region Protected Functions

        protected WorkBookType CheckWorkBookType(string fileName)
        {
            string extention = Path.GetExtension(fileName);
            if (extention.Equals("xlsx", StringComparison.OrdinalIgnoreCase))
                return WorkBookType.XLSX;
            if (extention.Equals("xls", StringComparison.OrdinalIgnoreCase))
                return WorkBookType.XLS;
            else
                return WorkBookType.Unknown;
        }

        #endregion


    }
}
