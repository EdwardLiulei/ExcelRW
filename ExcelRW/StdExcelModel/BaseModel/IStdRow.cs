using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelReadAndWrite.StdExcelModel.BaseModel
{
    public interface IStdRow:IStdSheetCompo
    {
        void SetHeight(int height);

        StdExcelCellBase GetCell(int rowNum);
    }
}
