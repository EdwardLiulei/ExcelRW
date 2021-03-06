﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;

namespace ExcelReadAndWrite.StdExcelModel.DataModel
{
    internal interface IDataSheetCompo
    {

        bool Bold{get;set;}

        bool Italic { set; get; }

        void SetFontStyle(Font font);

        void SetBackgroudColor(Color color);

        void SetFontColor(Color color);
    }
}
