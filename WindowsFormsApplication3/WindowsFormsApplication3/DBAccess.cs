using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient;
using System.Threading.Tasks;

namespace WindowsFormsApplication3
{
   public class DBAccess
    {
        private SqlConnection _sqlCon = null;
        private string _sqlConnectStr = null;
        private SqlCommand _sqlCmd = null;
        private SqlDataAdapter _sqlDA = null;
        private DataSet _ds = null;
        private SqlTransaction _sqlTran= null;
        StringBuilder _strBuilder = null;
       //构造函数初始化对象
        public DBAccess()
        {
            _sqlConnectStr = "server=shndtpdl0955\\SQL2008EXPR2;database=remedy_reports;uid=sa;pwd=password_1";
            _ds = new DataSet();
            _sqlCon = new SqlConnection(_sqlConnectStr);
            _sqlCmd = new SqlCommand();
            _sqlCmd.CommandTimeout = 40000;
            _sqlCmd.Connection = _sqlCon;
            _sqlDA = new SqlDataAdapter();
        }
       //Command执行方法体
        public void ExecuteCommand(String execSQLStr) {
            if (_sqlCon.State == ConnectionState.Open) {
                _sqlCon.Close();
            }
            _sqlCon.Open();
            _sqlCmd.CommandType = CommandType.Text;
            _sqlCmd.CommandText = execSQLStr;
            _sqlTran = _sqlCon.BeginTransaction();
            _sqlCmd.Transaction = _sqlTran;
            try
            {
                _sqlCmd.ExecuteNonQuery();
                _sqlTran.Commit();
            }
            catch
            {
                _sqlTran.Rollback();
            }
            finally {
                _sqlCon.Close();
            }
        }
       //向数据库中插入remedy7型的数据
        public void InsertDataToDB(DataSet dsSource, string targetTable) { 
            //先清空目标表里的数据
            this.ExecuteCommand("delete table " + targetTable);
            try
            {
                if (_sqlCon.State == ConnectionState.Open)
                {
                    _sqlCon.Close();
                }
                _sqlCon.Open();
                SqlTransaction stran1 = _sqlCon.BeginTransaction();
                _sqlCmd.Transaction = stran1;
                _sqlCmd.CommandType = CommandType.Text;
                DataTable dtSource = dsSource.Tables[0];
                foreach (DataRow dr in dtSource.Rows)
                {
                    _strBuilder = new StringBuilder();
                    _strBuilder.Append("Insert into " + targetTable + " values('");
                    for (int i = 0; i < dtSource.Columns.Count - 1; i++)
                    {
                        _strBuilder.Append(dr[i].ToString().Replace("'", "''") + "','");
                    }
                    _strBuilder.Append(dr[dtSource.Columns.Count - 1] + "')");
                    _sqlCmd.CommandText = _strBuilder.ToString();
                    _sqlCmd.ExecuteNonQuery();
                }
                stran1.Commit();
                if (_sqlCon.State == ConnectionState.Open) _sqlCon.Close();
            }
            catch (Exception E)
            {
                throw new ArgumentException("插入数据失败原因是"+E.Message);
            }

        }
       //插入Task Open数据
        public void InsertTaskData(DataSet taskdataset) {
            if (_sqlCon.State == ConnectionState.Open) {
                _sqlCon.Close();
            }
            _sqlCon.Open();
            _sqlTran = _sqlCon.BeginTransaction();
            _sqlCmd.Transaction = _sqlTran;
            bool noIdFlag = false;
            string lastIncID = taskdataset.Tables[0].Rows[6][0].ToString();
            string currIncID = lastIncID;
            for (int rowNum = 5; rowNum < taskdataset.Tables[0].Rows.Count - 1; rowNum++)//rowNum begin with 5, that depends on open task sheet.
            {
                _strBuilder = new StringBuilder();
                if (!noIdFlag)
                {
                    currIncID = taskdataset.Tables[0].Rows[rowNum][0].ToString();
                }
                _strBuilder.Append("Insert into TaskAssignedRaw values('" + currIncID + "','");
                for (int colNum = 1; colNum < taskdataset.Tables[0].Columns.Count - 2; colNum++)
                {
                    _strBuilder.Append(taskdataset.Tables[0].Rows[rowNum][colNum].ToString().Replace(",", "").Replace("'", "") + "','");
                }
                _strBuilder.Append("'," + "'')");

                _sqlCmd.CommandText = _strBuilder.ToString();
                try
                {
                    _sqlCmd.ExecuteNonQuery();
                }
                catch
                {
                    _sqlTran.Rollback();
                    throw new ArgumentException("Inserting raw data into table PSTDBRaw is failure, please check!");
                }
                //For merged cell, assign last valued cell
                if (string.IsNullOrEmpty(taskdataset.Tables[0].Rows[rowNum + 1][0].ToString()))
                {
                    lastIncID = currIncID;
                    noIdFlag = true;
                }
                else
                {
                    lastIncID = currIncID;
                    noIdFlag = false;
                }
            }

            _sqlTran.Commit();

            if (_sqlCon.State == ConnectionState.Open)
            {
                _sqlCon.Close();
            }
        }



        public void ExecuteProcWithParam(string fileName, string[] procNames, string[] sheetNames, string[,] parameters, bool isTwoParas, DateTime generateDate)
        {
            string para;
            for (int i = 0; i < procNames.Length; i++)
            {
                try
                {
                    _sqlCon.Open();
                    _sqlCmd.CommandType = CommandType.StoredProcedure;
                    string ProcName = procNames[i].ToString();
                    _sqlCmd.CommandText = ProcName;
                    if (ProcName == "TriBenchMarkRpt_MatchedBacklog")//add if condition for benchmark report.
                    {
                        DateTime EndDay = generateDate.AddDays(-2);
                        DateTime StartDay = generateDate.AddDays(-8);
                        string ParaEndDate = EndDay.ToString("yyyy-MM-dd");
                        string ParaStartDate = StartDay.ToString("yyyy-MM-dd");
                        _sqlCmd.Parameters.AddWithValue("@startDate", ParaStartDate);
                        _sqlCmd.Parameters.AddWithValue("@endDate", ParaEndDate);
                    }
                    else
                    {
                        if (isTwoParas)
                        {
                            for (int j = 0; j < 2; j++)
                            {
                                if (parameters[i, j] == "")
                                {
                                    ;
                                }
                                else
                                {
                                    if (j == 0)
                                    {
                                        para = "@diffn";
                                    }
                                    else
                                    {
                                        para = "@diffb";
                                    }
                                    _sqlCmd.Parameters.AddWithValue(para, Int32.Parse(parameters[i, j].ToString()));
                                }
                            }
                        }
                        else
                        {
                            for (int j = 0; j < 1; j++)
                            {
                                if (parameters[i, j] == "")
                                {
                                    ;
                                }
                                else
                                {
                                    if (j == 0)
                                    {
                                        para = "@diffn";
                                    }
                                    else
                                    {
                                        para = "@diffb";
                                    }
                                    _sqlCmd.Parameters.AddWithValue(para, Int32.Parse(parameters[i, j].ToString()));
                                }
                            }
                        }
                    }

                    _sqlDA.SelectCommand = _sqlCmd;
                    _sqlDA.Fill(_ds, ProcName);
                }
                catch (Exception e)
                {
                    throw e;
                }
                finally
                {
                    _sqlCmd.Parameters.Clear();
                    _sqlCon.Close();
                }
            }
            //filename1:the template file name
            ExcelBus excelBiz = new ExcelBus();
            excelBiz.ExportDataToExcel(fileName, procNames, sheetNames, _ds, generateDate);
        }


        public DataSet getDSByProcs(string[] ProcNames, string[,] ParamValues)
        {
            //with params
            try
            {
                _sqlCon.Open();
                _sqlCmd.CommandType = CommandType.StoredProcedure;
                for (int i = 0; i < ProcNames.Length; i++)
                {
                    string execProcName = ProcNames[i].ToString();//get execute proc name 
                    if (string.IsNullOrEmpty(ParamValues[i, 2].ToString()))
                    {
                        _sqlCmd.CommandText = execProcName;
                        _sqlDA.SelectCommand = _sqlCmd;
                        _sqlDA.Fill(_ds, ProcNames[i].ToString());
                    }
                    else
                    {
                        DateTime startDateValue = Convert.ToDateTime(ParamValues[i, 0]);//get execute proc parameter value
                        DateTime endtDateValue = Convert.ToDateTime(ParamValues[i, 1]);//get execute proc parameter value
                        int isMultiDay = Int32.Parse(ParamValues[i, 2]);//get execute proc parameter value
                        _sqlCmd.CommandText = execProcName;
                        _sqlCmd.Parameters.AddWithValue("@startdate", startDateValue);
                        _sqlCmd.Parameters.AddWithValue("@enddate", endtDateValue);
                        _sqlCmd.Parameters.AddWithValue("@type", isMultiDay);//if value of isMultiDay is 0, means one day report
                        _sqlDA.SelectCommand = _sqlCmd;
                        _sqlDA.Fill(_ds, ProcNames[i].ToString());
                    }
                    _sqlCmd.Parameters.Clear();
                }
            }
            catch (Exception e)
            {
                throw e;
            }
            finally
            {
                _sqlCmd.Parameters.Clear();
                _sqlCon.Close();
            }
            return _ds;

        }


        public DataSet ExecuteProcWithParam_Navy(string fileName, string[] procNames, string[] sheetNames, string[,] parameters, DateTime generateDate)
        {
            string para;
            for (int i = 0; i < procNames.Length; i++)
            {
                try
                {
                    _sqlCon.Open();
                    _sqlCmd.CommandType = CommandType.StoredProcedure;
                    string ProcName = procNames[i].ToString();
                    _sqlCmd.CommandText = ProcName;
                    for (int j = 0; j < 2; j++)
                    {
                        if (j == 0)
                        {
                            para = "@report_start_date";
                        }
                        else
                        {
                            para = "@report_end_date";
                        }
                        _sqlCmd.Parameters.AddWithValue(para, parameters[i, j]);
                    }
                    _sqlDA.SelectCommand = _sqlCmd;
                    _sqlDA.Fill(_ds, ProcName);
                }
                catch (Exception e)
                {
                    throw e;
                }
                finally
                {
                    _sqlCmd.Parameters.Clear();
                    _sqlCon.Close();
                }
            }
            //filename1:the template file name
            //ExcelBusiness excelBiz = new ExcelBusiness();
            //excelBiz.ExportDataToExcel_Navy(fileName, procNames, sheetNames, _ds, generateDate);
            return _ds;
        }

    }
}
