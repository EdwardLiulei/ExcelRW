using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;
using System.Drawing.Text;

namespace ExcelReadAndWrite.StdExcelModel.BaseModel
{
    public interface IStdCell
    {
        string GetValue();

        bool IsMerged();

        void SetValue(string value);

        void SetFormular(string formular);

        void SetFontStyle(Font font);

        void SetBold();

        void SetItalic();

        void SetBackgroudColor(Color color);

        void SetFontColor(Color color);
    }
}
