﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelReadAndWrite.StdExcelModel.DataModel;
using System.Drawing;

namespace ExcelReadAndWrite.StdExcelModel
{
    public abstract class StdExcelCellBase
    {

        #region Field
        private bool _isMerged;
        #endregion

        #region Properity

        public abstract bool Bold { set; get; }

        public abstract bool Italic { set; get; }
        #endregion

        #region Abstract Functions
        public abstract string GetValue();

        public abstract void SetValue(string value);

        public abstract void SetFormular(string formular);

        public abstract void SetFontStyle(Font font);


        public abstract void SetBackgroudColor(Color color);

        public abstract void SetFontColor(Color color);

        #endregion

        public bool IsMerged()
        {
            return _isMerged;
        }
    }
}
