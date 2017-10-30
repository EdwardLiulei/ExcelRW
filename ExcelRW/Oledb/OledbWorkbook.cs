using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelReadAndWrite.StdExcelModel;
using System.Data.OleDb;
using System.Data;

namespace ExcelReadAndWrite.Oledb
{
    public class OledbWorkbook:StdExcelWorkbookBase
    {
        #region Field
        private string _connectStr;
        private DataSet _dataset;
        #endregion

        #region Constuctor

        public OledbWorkbook() : base()
        {

        }

        public OledbWorkbook(string fileName):base(fileName)
        {

        }

        #endregion

        protected override void ReadWorkbook(string fileName)
        {
            string provider = "Provider=Microsoft.ACE.OLEDB.12.0;";
            string dataSource = "Data Source="+fileName+";";
            string extended = "Extended Properties='Excel 8.0;HDR=NO;IMEX=1'";

            _connectStr = provider + dataSource + extended;
            OleDbConnection objConn = new OleDbConnection(_connectStr);
            objConn.Open();
            System.Data.DataTable dt = objConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
            _dataset = new DataSet();
            

            foreach (System.Data.DataRow row in dt.Rows)
            {

                string sheetName = row["TABLE_NAME"].ToString().Replace("$","");

                string strExcel = "select * from [" + sheetName + "$]";
                var myCommand = new OleDbDataAdapter(strExcel, _connectStr);

                myCommand.Fill(_dataset, sheetName);

                _workSheets.Add(new OledbWorksheet(_dataset.Tables[sheetName], sheetName, _connectStr));
            }
            //throw new NotImplementedException();
        }

        public override void Save(string fileName)
        {
            var sConn = string.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};" + "Extended Properties='Excel 12.0 XML;HDR=NO'", fileName);
            _connectStr = sConn;
            OleDbConnection myConn = new OleDbConnection(sConn);

            string sqlCreate = "CREATE TABLE TestSheet";

            OleDbCommand cmd = new OleDbCommand(sqlCreate, myConn);

            //创建Excel文件：C:/test.xls

            myConn.Open();

            //创建TestSheet工作表

            cmd.ExecuteNonQuery();

            //添加数据

            cmd.CommandText = "INSERT INTO TestSheet VALUES(1,'elmer','password')";

            cmd.ExecuteNonQuery();

            //关闭连接

            myConn.Close();

            //foreach (var worksheet in _workSheets)
            //{
            //    ExcuteSQL(_dataset, worksheet.SheetName, _connectStr);
            //}
        }

        private static void ExcuteSQL(DataSet oldds, string tableName, string strCon)
        {
            //连接
            OleDbConnection myConn = new OleDbConnection(strCon);

            string strCom = "select * from [" + tableName + "$]";

            try
            {
                myConn.Open();

                OleDbDataAdapter myCommand = new OleDbDataAdapter(strCom, myConn);

                System.Data.OleDb.OleDbCommandBuilder builder = new OleDbCommandBuilder(myCommand);

                //QuotePrefix和QuoteSuffix主要是对builder生成InsertComment命令时使用。   
                //获取insert语句中保留字符（起始位置）  
                builder.QuotePrefix = "[";

                //获取insert语句中保留字符（结束位置）   
                builder.QuoteSuffix = "]";

                DataSet newds = new DataSet();
                //获得表结构
                DataTable ndt = oldds.Tables[0].Clone();
                //清空数据
                //ndt.Rows.Clear();

                ndt.TableName = tableName;
                newds.Tables.Add(ndt);

                //myCommand.Fill(newds, TableName);

                for (int i = 0; i < oldds.Tables[0].Rows.Count; i++)
                {
                    //在这里不能使用ImportRow方法将一行导入到news中，
                    //因为ImportRow将保留原来DataRow的所有设置(DataRowState状态不变)。
                    //在使用ImportRow后newds内有值，但不能更新到Excel中因为所有导入行的DataRowState!=Added     
                    DataRow nrow = newds.Tables[0].NewRow();
                    for (int j = 0; j < oldds.Tables[0].Columns.Count; j++)
                    {
                        nrow[j] = oldds.Tables[0].Rows[i][j];
                    }
                    newds.Tables[0].Rows.Add(nrow);
                }

                myCommand.Update(newds, tableName);
            }
            finally
            {
                myConn.Close();
            }
        }



        public override StdExcelWorkSheetBase GetSheet(string sheetName)
        {
            throw new NotImplementedException();
        }

        public override StdExcelWorkSheetBase InsertSheet(string sheetName)
        {
            throw new NotImplementedException();
        }

        public override StdExcelWorkSheetBase InsertSheet(DataTable table, string sheetName, bool withHeader)
        {
            throw new NotImplementedException();
        }
    }
}
