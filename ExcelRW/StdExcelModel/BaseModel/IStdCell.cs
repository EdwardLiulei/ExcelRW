using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;
using System.Drawing.Text;

namespace ExcelReadAndWrite.StdExcelModel.BaseModel
{
    public interface IStdCell:IStdSheetCompo
    {
        string GetValue();

        bool IsMerged();

        void SetValue(string value);

        void SetFormular(string formular);
    }
}
