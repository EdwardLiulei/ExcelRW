using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelReadAndWrite.StdExcelModel.BaseModel;

namespace ExcelReadAndWrite.StdExcelModel
{
    public abstract class StdExcelWorkbookBase:IStdWorkbook
    {
        #region Filed
        protected List<StdExcelWorkSheetBase> _workSheets;

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
        #endregion

        public abstract StdExcelWorkSheetBase GetSheet(string sheetName);

        

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
        


        
    }
}
