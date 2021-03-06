﻿using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using ExcelReadAndWrite.StdExcelModel;
using System.Drawing;
using Microsoft.Office.Interop.Excel;

namespace ExcelReadAndWrite.Com
{
    public class ComExcelRange : StdExcelRangeBase
    {

        #region Field
        private Range _comRange;
        #endregion


        #region Properity

        public override bool Bold
        {
            set
            { _comRange.Font.Bold = value; }
            get
            { return _comRange.Font.Bold; }
        }

        public override bool Italic
        {
            set
            { _comRange.Font.Italic = value; }
            get
            { return _comRange.Font.Italic; }
        }

        #endregion

        #region Constructor
        public ComExcelRange(Range range)
        {
            _comRange = range;
 
        }

        #endregion
        public override void SetBackgroudColor(Color color)
        {
            _comRange.Interior.Color = System.Drawing.Color.FromArgb(color.A, color.B, color.G, color.R).ToArgb(); 
        }

      

        public override void SetFontColor(Color color)
        {
            _comRange.Font.Color = color;
        }

        public override void SetFontStyle(System.Drawing.Font font)
        {

            _comRange.Font.Name = font.Name;
            _comRange.Font.Size = font.Size;
            _comRange.Font.Bold = font.Bold;
            _comRange.Font.Italic = font.Italic;
            _comRange.Font.Underline = font.Underline;
        }



        public override void SetMerge()
        {
            _comRange.MergeCells = true;
        }

        public override void UnMerge()
        {
            _comRange.MergeCells = false;
        }

        public override string[,] GetRangeData()
        {
            string [,] value = _comRange.Value2;
            return value;
        }
    }
}
