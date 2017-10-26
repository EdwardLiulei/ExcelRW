using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelReadAndWrite.StdExcelModel.BaseModel
{
    internal interface IStdRow:IStdSheetCompo
    {
        void SetHeight(int height);

        StdExcelCellBase GetCell(int rowNum);
    }
}
