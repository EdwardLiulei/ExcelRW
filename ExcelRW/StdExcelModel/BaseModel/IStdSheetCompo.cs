using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;

namespace ExcelReadAndWrite.StdExcelModel.BaseModel
{
    public interface IStdSheetCompo
    {



        void SetFontStyle(Font font);

        void SetBold();

        void SetItalic();

        void SetBackgroudColor(Color color);

        void SetFontColor(Color color);
    }
}
