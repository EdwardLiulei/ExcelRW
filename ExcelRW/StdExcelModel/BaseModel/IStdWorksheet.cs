using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;
using System.Data;

namespace ExcelReadAndWrite.StdExcelModel.BaseModel
{
    public interface IStdWorksheet
    {
        
        string GetSheetName();

        IStdRange GetRange(int startRow, int startCol, int endRow, int endCol);

        IStdCell GetCell(int rowNum, int columnNum);

        string GetCellValue(int rowNum, int columnNum);

        string GetCellFormular(int rowNum, int columnNum);

        IStdRow GetRow(int index);

        IStdColumn GetColumn(int index);

        void InsertRow(int index);

        void InsertColumn(int index);


        void SetCellValue(string value, int rowNum, int columnNum);

        void SetCellFormular(string formular, int rowNum, int columnNum);

        void SetRangeColor(IStdRange range,Color color);
        void SetCellColor(int rowNum, int columnNum, Color color);

        void MergeCell(IStdRange range);
        void MergeCell(int startRow, int startCol, int endRow, int endCol);

        DataTable GetTableContent();
    }
}
