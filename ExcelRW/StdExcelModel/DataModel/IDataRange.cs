using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelReadAndWrite.StdExcelModel.DataModel
{
    internal interface IDataRange:IDataSheetCompo
    {

        void SetMerge();
        void UnMerge();
    }
}
