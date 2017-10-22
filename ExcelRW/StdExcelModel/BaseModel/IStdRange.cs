using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelReadAndWrite.StdExcelModel.BaseModel
{
    public interface IStdRange:IStdSheetCompo
    {

        void SetMerge();
        void UnMerge();
    }
}
