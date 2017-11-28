using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelReadAndWrite.StdExcelModel.DataModel
{
    internal interface IDataColumn:IDataSheetCompo
    {
        void SetWidth(int width);

        IDataCell GetCell(int columnNum);
    }
}
