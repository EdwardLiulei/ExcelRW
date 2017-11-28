using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;
using System.Data;

namespace ExcelReadAndWrite.StdExcelModel.DataModel
{
    public interface IDataWorksheet<out IDataRange,out IDataCell>
    {
        
        string GetSheetName();

        IDataRange GetRange(int startRow, int startCol, int endRow, int endCol);

        IDataCell GetCell(int rowNum, int columnNum);

        string GetCellValue(int rowNum, int columnNum);

        string GetCellFormular(int rowNum, int columnNum);

        StdExcelRowBase GetRow(int index);

        StdExcelColumnBase GetColumn(int index);

        void InsertRow(int index);

        void InsertColumn(int index);


        void SetCellValue(string value, int rowNum, int columnNum);

        void SetCellFormular(string formular, int rowNum, int columnNum);



        DataTable GetTableContent();
    }
}
