using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelReadAndWrite.StdExcelModel.BaseModel
{
    public interface IStdWorkbook
    {
        List<StdExcelWorkSheetBase>WorkSheets { get; }
       // void ReadWorkbook(string fileName);
        void Save(string fileName);

        List<string> GetSheetList();
        int GetSheetCount();
        StdExcelWorkSheetBase GetSheetByName(string sheetName);
        StdExcelWorkSheetBase GetSheetByIndex(int index);

        string GetSheetNameByIndex(int index);

        StdExcelWorkSheetBase CloneSheet(int index);
        StdExcelWorkSheetBase CloneSheet(string sheetName);

        bool Is1904();

    }
}
