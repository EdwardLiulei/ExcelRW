using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelReadAndWrite.StdExcelModel.DataModel
{
    internal interface IDataRow:IDataSheetCompo
    {
        void SetHeight(int height);

        StdExcelCellBase GetCell(int rowNum);
    }
}
