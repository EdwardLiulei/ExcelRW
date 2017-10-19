using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelReadAndWrite.StdExcelModel.BaseModel;

namespace ExcelReadAndWrite.StdExcelModel
{
    public abstract class ExcelWorkbookBase:IStdWorkbook
    {
        #region Filed
        protected List<ExcelWorkSheetBase> _workSheets;

        #endregion

        #region Properity

        public List<ExcelWorkSheetBase> WorkSheets { get { return _workSheets; } }
        #endregion

        #region Constructor
        public ExcelWorkbookBase()
        {
            _workSheets = new List<ExcelWorkSheetBase>();
        }
        #endregion

        #region Absrtract Fuctions

        public abstract void Load(string fileName);

        public abstract void Save(string fileName);
        #endregion

        public abstract ExcelWorkSheetBase GetSheet(string sheetName);

        

        public List<string> GetSheetList()
        {
            return _workSheets.Select(p => p.GetSheetName()).ToList();
        }

        public int GetSheetCount()
        {
            return _workSheets.Count();
        }

        public ExcelWorkSheetBase GetSheetByName(string sheetName)
        {
            return _workSheets.Find(p => p.GetSheetName().Equals(sheetName,StringComparison.OrdinalIgnoreCase));
        }

        public ExcelWorkSheetBase GetSheetByIndex(int index)
        {
            return _workSheets[index];
        }

        public string GetSheetNameByIndex(int index)
        {
            return _workSheets[index].GetSheetName();
        }

        public ExcelWorkSheetBase CloneSheet(int index)
        {
            throw new NotImplementedException();
        }

        public ExcelWorkSheetBase CloneSheet(string sheetName)
        {
            throw new NotImplementedException();
        }

        public bool Is1904()
        {
            throw new NotImplementedException();
        }
        


        
    }
}
