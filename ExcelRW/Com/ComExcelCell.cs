﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelReadAndWrite.StdExcelModel;
using System.Drawing;
using Microsoft.Office.Interop.Excel;

namespace ExcelReadAndWrite.Com
{
    public class ComExcelCell:StdExcelCellBase
    {
        #region Field
        private Range _comCell;
        
        #endregion
        #region Constructor
        public ComExcelCell(Range cell)
        {
            _comCell = cell;
 
        }

        #endregion
        public override string GetValue()
        {
            object o = _comCell.Value;
            return o.ToString();
        }

        public override void SetValue(string value)
        {
            _comCell.Value = value;
        }

        public override void SetFormular(string formular)
        {
            _comCell.Formula = formular;
        }

        public override void SetFontStyle(System.Drawing.Font font) 
        {
           
            _comCell.Font.Name =  font.Name;
            _comCell.Font.Size = font.Size;
            _comCell.Font.Bold = font.Bold;
            _comCell.Font.Italic = font.Italic;
            _comCell.Font.Underline = font.Underline;
        }

        public override void SetBold() 
        {
            _comCell.Font.Bold = true;
        }

        public override void SetItalic() 
        {
            _comCell.Font.Italic = true;
        }

        public override void UnBold()
        {
            _comCell.Font.Bold = false;
        }

        public override void UnItalic()
        {
            _comCell.Font.Italic = false;
        }

        public override void SetBackgroudColor(Color color) 
        {
            
            _comCell.Interior.Color = System.Drawing.Color.FromArgb(color.A, color.B, color.G, color.R).ToArgb();
        }

        public override void SetFontColor(Color color)
        {
            _comCell.Font.Color = System.Drawing.Color.FromArgb(color.A, color.B, color.G, color.R).ToArgb();
        }
    }
}
