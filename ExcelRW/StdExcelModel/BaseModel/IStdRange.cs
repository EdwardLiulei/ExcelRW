using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelReadAndWrite.StdExcelModel.BaseModel
{
    internal interface IStdRange:IStdSheetCompo
    {

        void SetMerge();
        void UnMerge();
    }
}
