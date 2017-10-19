using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelReadAndWrite.StdExcelModel.BaseModel
{
    public interface IStdWorkbook
    {
        List<ExcelWorkSheetBase>WorkSheets { get; }
        void Load(string fileName);
        void Save(string fileName);

        List<string> GetSheetList();
        int GetSheetCount();
        ExcelWorkSheetBase GetSheetByName(string sheetName);
        ExcelWorkSheetBase GetSheetByIndex(int index);

        string GetSheetNameByIndex(int index);

        ExcelWorkSheetBase CloneSheet(int index);
        ExcelWorkSheetBase CloneSheet(string sheetName);

        bool Is1904();

    }
}
