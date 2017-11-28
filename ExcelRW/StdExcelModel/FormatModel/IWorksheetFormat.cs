using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;

namespace ExcelReadAndWrite.StdExcelModel.FormatModel
{
    public interface IWorksheetFormat
    {
        void SetRangeColor(StdExcelRangeBase range, Color color);
        void SetCellColor(int rowNum, int columnNum, Color color);

        void MergeCell(StdExcelRangeBase range);
        void MergeCell(int startRow, int startCol, int endRow, int endCol);
    }
}
