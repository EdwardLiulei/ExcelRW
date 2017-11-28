using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelReadAndWrite.StdExcelModel.FormatModel
{
    public interface ICellFormat
    {
        bool IsMerged();

        void SetValue(string value);

        void SetFormular(string formular);
    }
}
