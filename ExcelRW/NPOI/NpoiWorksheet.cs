﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelReadAndWrite.StdExcelModel;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using  DataTable = System.Data.DataTable;
using System.Data;

namespace ExcelReadAndWrite.NPOI
{
    public class NpoiWorksheet:ExcelWorkSheetBase
    {
        private ISheet _npoiWorsheet;

        public NpoiWorksheet(ISheet worksheet)
        {
            _npoiWorsheet = worksheet;
            _sheetName = worksheet.SheetName;
        }

        public override string GetCellValue(int rowNumber, int columNumber)
        {
            throw new NotImplementedException();
        }

        public override System.Data.DataTable GetTableContent()
        {
            bool isColumnName = true;
            int startRow = 1;
            ICell cell;
            DataColumn column;
            IRow row;
            System.Data.DataTable dataTable = new DataTable();
            if (_npoiWorsheet == null)
                return dataTable;

            int rowCount = _npoiWorsheet.LastRowNum;//总行数  
            if (rowCount > 0)
            {
                IRow firstRow = _npoiWorsheet.GetRow(0);//第一行  
                int cellCount = firstRow.LastCellNum;//列数  

                //构建datatable的列  
                if (isColumnName)
                {
                    startRow = 1;//如果第一行是列名，则从第二行开始读取  
                    for (int i = firstRow.FirstCellNum; i < cellCount; ++i)
                    {
                        cell = firstRow.GetCell(i);
                        if (cell != null)
                        {
                            string cellValue = GetCellValue(cell).ToString();
                            column = new DataColumn(cellValue);
                            dataTable.Columns.Add(column);

                        }
                    }
                }
                else
                {
                    for (int i = firstRow.FirstCellNum; i < cellCount; ++i)
                    {
                        column = new DataColumn("column" + (i + 1));
                        dataTable.Columns.Add(column);
                    }
                }

                //填充行  
                for (int i = startRow; i <= rowCount; ++i)
                {
                    row = _npoiWorsheet.GetRow(i);
                    if (row == null) continue;

                    DataRow dataRow = dataTable.NewRow();
                    for (int j = row.FirstCellNum; j < cellCount; ++j)
                    {
                        cell = row.GetCell(j);
                        if (cell == null)
                        {
                            dataRow[j] = "";
                        }
                        else
                        {
                            dataRow[j] = GetCellValue(cell);
                        }
                    }
                    dataTable.Rows.Add(dataRow);
                }
            }



            return dataTable;
        }

        public object GetCellValue(ICell cell)
        {
            if (cell == null)
                return "";
            object value = null;
            try
            {
                if (cell.CellType != CellType.Blank)
                {
                    switch (cell.CellType)
                    {
                        case CellType.Numeric:
                            // Date comes here
                            if (DateUtil.IsCellDateFormatted(cell))
                            {
                                value = cell.DateCellValue;
                            }
                            else
                            {
                                // Numeric type
                                value = cell.NumericCellValue;
                            }
                            break;
                        case CellType.Boolean:
                            // Boolean type
                            value = cell.BooleanCellValue;
                            break;
                        case CellType.Formula:
                            value = cell.CellFormula;
                            break;
                        default:
                            // String type
                            value = cell.StringCellValue;
                            break;
                    }
                }
            }
            catch (Exception)
            {
                value = "";
            }

            return value;
        }
    }
}