using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Runtime.InteropServices;
using System.Data.OleDb;
using Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.IO;
namespace WindowsFormsApplication3
{
  public  class ExcelBus
  {
      private static string packagePath = System.IO.Directory.GetCurrentDirectory().ToString().Replace(@"bin\Debug", @"SSISPackage\");//获取SSISPackage的目录。
      private static string excelTemplatePath = System.IO.Directory.GetCurrentDirectory().ToString().Replace(@"bin\Debug", @"ExcelTemplate\");//获取Exceltemplate的目录。
      private static string finalReportsPath = "D:\\PSTWork\\AutoReports\\Final reports\\";
      private static string trackWeekly4Path = "D:\\PSTWork\\Track Weekly(Thursday)\\";
      private static string trackWeekly1Path = "D:\\PSTWork\\Track Weekly(Monday)\\";
      private static string corpTrackWeekly1Path = "D:\\PSTWork\\Track Corp Weekly(Monday)\\";

      private string _fullFilePath;
      private int _IniRow;
      private string _sheetName;
      private string _SName;

      private _Workbook _ExcelWBook;
      private _Worksheet _ExcelWSheet;
      private Microsoft.Office.Interop.Excel.Application _ExcelApp;
      //获取Sheet Name
      public string GetFirstSheetName(string filePath)
      {
          
          string sheetName = string.Empty;
          try
          {
              _ExcelApp = new Microsoft.Office.Interop.Excel.Application();
              _ExcelWBook = (_Workbook)(_ExcelApp.Workbooks.Open(filePath, Missing.Value, Missing.Value
                          , Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                          Missing.Value, Missing.Value, Missing.Value));
              _ExcelWSheet = (Worksheet)_ExcelWBook.Sheets[1];
              sheetName = _ExcelWSheet.Name;

              _ExcelApp.DisplayAlerts = false;
              _ExcelWBook.Close(Missing.Value, Missing.Value, Missing.Value);
              _ExcelApp.Workbooks.Close();
              _ExcelApp.Quit();

              return sheetName;
          }
          catch (Exception)
          {
              _ExcelWBook.Close(Missing.Value, Missing.Value, Missing.Value);
              _ExcelApp.Workbooks.Close();
              _ExcelApp.Quit();
              return sheetName;
          }

      }
      public DataSet getDataSetFromExcel(string excelPath,string excelTable) {
          string olestr = string.Format("Provider = Microsoft.ACE.OLEDB.12.0;Data Source = {0};Extended Properties=\"Excel 12.0 Xml;HDR=YES\"", excelPath);
          OleDbConnection oconn = new OleDbConnection(olestr);
          OleDbCommand ocmd = new OleDbCommand();
          OleDbDataAdapter oda = new OleDbDataAdapter();
          ocmd.Connection = oconn;
          ocmd.CommandText = "select * from ["+excelTable+"]";
          oda.SelectCommand = ocmd;
          DataSet dset = new DataSet();
          try
          {
              oconn.Open();
              oda.Fill(dset);
              return dset;
          }
          catch (Exception e)
          {
              throw e;
          }
          finally {
              oconn.Close();
          }
      }

      public void PreExitExcel()
      {
          System.Diagnostics.Process[] allProcess = System.Diagnostics.Process.GetProcesses();
          foreach (System.Diagnostics.Process thisprocess in allProcess)
          {
              string processName = thisprocess.ProcessName;
              if (processName.ToLower() == "excel")
              {
                  try
                  {
                      thisprocess.Kill();
                  }
                  finally
                  {
                      ;
                  }
              }
          }
      }

      public void ExportDataToExcel(string fileName, string[] tableName, string[] sheetName, DataSet ds, DateTime generateDate)
      {
          PreExitExcel();
          string filename = fileName.Substring(0, fileName.LastIndexOf("."));
          try
          {
              _ExcelApp = new Microsoft.Office.Interop.Excel.Application();
              _ExcelWBook = (_Workbook)(_ExcelApp.Workbooks.Open(excelTemplatePath + filename, Missing.Value, Missing.Value
                      , Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                      Missing.Value, Missing.Value, Missing.Value));

              for (int i = 0; i < tableName.Length; i++)
              {
                  if (tableName[i] == "TriBenchMarkRpt_MatchedBacklog")
                  {
                      _IniRow = 2;
                      _sheetName = sheetName[i].ToString();
                      _ExcelWSheet = (_Worksheet)_ExcelWBook.Sheets[sheetName[i].ToString()];
                      for (int j = 0; j < ds.Tables[tableName[i].ToString()].Rows.Count; j++)
                      {
                          for (int k = 0; k < ds.Tables[tableName[i].ToString()].Columns.Count; k++)
                          {
                              _ExcelWSheet.Cells[_IniRow, k + 2] = ds.Tables[tableName[i].ToString()].Rows[j][k].ToString();
                          }
                          _IniRow++;
                      }
                  }
                  else
                  {
                      _IniRow = 3;
                      _sheetName = sheetName[i].ToString();
                      _ExcelWSheet = (_Worksheet)_ExcelWBook.Sheets[sheetName[i].ToString()];
                      int rowcount = ds.Tables[i].Rows.Count;
                      for (int j = 0; j < ds.Tables[tableName[i].ToString()].Rows.Count; j++)
                      {
                          for (int k = 0; k < ds.Tables[tableName[i].ToString()].Columns.Count; k++)
                          {
                              _ExcelWSheet.Cells[_IniRow, k + 1] = ds.Tables[tableName[i].ToString()].Rows[j][k].ToString();
                          }
                          _IniRow++;
                      }
                  }
              }

              string strDate = generateDate.AddDays(-1).ToString("MMdd");
              ExcelHelper excelHelper = new ExcelHelper();
              _fullFilePath = string.Format(excelHelper.getFilename(generateDate) + "\\{0}-{1}.xls", filename, strDate);
              _ExcelApp.Rows.RowHeight = "15";

              //ExApp.DisplayAlerts = false;
              _ExcelWBook.CheckCompatibility = false;//Add for Diable Compatibility for saving excel
              _ExcelWBook.SaveAs(_fullFilePath, XlFileFormat.xlWorkbookNormal,
                  null, null, false, false, XlSaveAsAccessMode.xlExclusive, false, false, null, null, null);
              _ExcelWBook.Close(Missing.Value, Missing.Value, Missing.Value);
              _ExcelApp.Workbooks.Close();
              _ExcelApp.Quit();

              Marshal.ReleaseComObject(_ExcelWSheet);
              Marshal.ReleaseComObject(_ExcelWBook);
              Marshal.ReleaseComObject(_ExcelApp);
          }
          catch (Exception e)
          {
              _ExcelWBook.Close(Missing.Value, Missing.Value, Missing.Value);
              _ExcelApp.Workbooks.Close();
              _ExcelApp.Quit();
              throw e;
          }
      }


      public void AEExportDataToExcel(string fileName, string[] tableNames, string[] sheetNames, DataSet ds, int[] irow, DateTime generateDate)
      {
          PreExitExcel();
          string filename1 = fileName.Substring(0, fileName.LastIndexOf("."));
          try
          {
              _ExcelApp = new Microsoft.Office.Interop.Excel.Application();
              _ExcelWBook = (_Workbook)(_ExcelApp.Workbooks.Open(excelTemplatePath + filename1, Missing.Value, Missing.Value
                      , Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                      Missing.Value, Missing.Value, Missing.Value));
              for (int i = 0; i < tableNames.Length; i++)
              {
                  _IniRow = irow[i];
                  switch (tableNames[i].ToString())
                  {
                      case "AnnuityOpenPeriod":
                      case "AnnuityUrgent":
                      case "AnnuityStatusUrgent":
                          _SName = "AnnuityFinal";
                          break;
                      case "AnnuityRpt":
                          _SName = "AnnuityRpt";
                          break;
                      case "AnnuityStatus":
                          _SName = "AnnuityStatus";
                          break;
                      case "EBusinessRpt":
                          _SName = "EBusinessRpt Grouped";
                          break;
                      default: ;
                          break;
                  }
                  _ExcelWSheet = (_Worksheet)_ExcelWBook.Sheets[_SName];
                  for (int j = 0; j < ds.Tables[tableNames[i]].Rows.Count; j++)
                  {
                      for (int k = 0; k < ds.Tables[tableNames[i]].Columns.Count; k++)
                      {
                          _ExcelWSheet.Cells[_IniRow, k + 1] = ds.Tables[tableNames[i]].Rows[j][k].ToString();
                      }
                      _IniRow++;
                  }
              }
              string strDate = generateDate.AddDays(-1).ToString("MMdd");
              ExcelHelper excelHelper = new ExcelHelper();
              _fullFilePath = string.Format(excelHelper.getFilename(generateDate) + "\\{0}-{1}.xls", filename1, strDate);
              _ExcelApp.Rows.RowHeight = "15";
              _ExcelWBook.SaveAs(_fullFilePath, XlFileFormat.xlWorkbookNormal,
                  null, null, false, false, XlSaveAsAccessMode.xlExclusive, false, false, null, null, null);
              _ExcelApp.DisplayAlerts = false;
              _ExcelWBook.Close(Missing.Value, Missing.Value, Missing.Value);
              _ExcelApp.Workbooks.Close();
              _ExcelApp.Quit();
              Marshal.ReleaseComObject(_ExcelWSheet);
              Marshal.ReleaseComObject(_ExcelWBook);
              Marshal.ReleaseComObject(_ExcelApp);
          }
          catch (Exception e)
          {
              _ExcelWBook.Close(Missing.Value, Missing.Value, Missing.Value);
              _ExcelApp.Workbooks.Close();
              _ExcelApp.Quit();
              throw e;
          }
      }


      public void PrepareFilesForDailyReport(DateTime generateDate)
      {
          String dweek = generateDate.DayOfWeek.ToString();
          DateTime priorWeekDate = generateDate.AddDays(-8);
          DateTime priorWeekDateS = generateDate.AddDays(-10);
          DateTime FridayDate = generateDate.AddDays(-3);
          DateTime dTodayLinkDate = generateDate.AddDays(-1);

          DateTime rTodayLinkDate = dweek == "Monday" ? generateDate.AddDays(-4) : generateDate.AddDays(-2);

          string tldDate = dTodayLinkDate.ToString("yyyyMMdd");
          string tlrDate = rTodayLinkDate.ToString("yyyyMMdd");
          string pWeekDate = priorWeekDate.ToString("yyyyMMdd");
          string rDatey = rTodayLinkDate.Year.ToString();
          string rDatem = rTodayLinkDate.ToString("MM");
          string rDated = rTodayLinkDate.ToString("dd");
          string dDatey = dTodayLinkDate.Year.ToString();
          string dDatem = dTodayLinkDate.ToString("MM");
          string dDated = dTodayLinkDate.ToString("dd");

          ExcelHelper excelHelper = new ExcelHelper();

          //export today link 
          string ssisPathTodaylink = packagePath + "todaylink.dtsx";
          string excelFullNameTodaylink = excelHelper.getFilenamePF(generateDate) + @"\" + dDatem + dDated + dDatey + "(Exclude Closed Before 6302009).xls";
          excelHelper.ExportDataBySSIS(ssisPathTodaylink, excelFullNameTodaylink);

          //export Tricare dump
          string ssisPathTricareDaily = packagePath + "tricaredaily.dtsx";
          string excelFullNameTricareDaily = excelHelper.getFilenamePF(generateDate) + @"\" + "Track Wise Dump-" + dDatem + dDated + dDatey + ".xls";
          excelHelper.ExportDataBySSIS(ssisPathTricareDaily, excelFullNameTricareDaily);

          //Annuity and Ebusiness report
          File.Copy(finalReportsPath + tldDate + @"\" + "Annuity and EBusiness-" + dDatem + dDated + ".xls", excelHelper.getFilenamePF(generateDate) + @"\" + "Annuity and EBusiness-" + dDatem + dDated + ".xls", true);

          ////business surpport backlog report
          File.Copy(finalReportsPath + tldDate + @"\" + "Service and TricareDaily-" + dDatem + dDated + ".xls", excelHelper.getFilenamePF(generateDate) + @"\" + "Service and TricareDaily-" + dDatem + dDated + ".xls", true);

          ////service remedy group level report
          File.Copy(finalReportsPath + tldDate + @"\" + "Service Remedy Group Level Report-" + dDatem + dDated + ".xls", excelHelper.getFilenamePF(generateDate) + @"\" + "Service Remedy Group Level Report-" + dDatem + dDated + ".xls", true);

          ////ebusiness dump
          File.Copy(finalReportsPath + tldDate + @"\" + "EBusiness ticket-" + dDatem + dDated + ".xls", excelHelper.getFilenamePF(generateDate) + @"\" + "EBusiness ticket-" + dDatem + dDated + dDatey + ".xls", true);

          ////annuity backlog dump
          File.Copy(finalReportsPath + tldDate + @"\" + "Annuity ticket-" + dDatem + dDated + ".xls", excelHelper.getFilenamePF(generateDate) + @"\" + "Annuity ticket-" + dDatem + dDated + dDatey + ".xls", true);

          System.Diagnostics.Process.Start("explorer.exe", excelHelper.getFilenamePF(generateDate));
      }



      public void PrepareFilesForTricareRpt(DateTime generateDate)
      {
          DateTime endDay = generateDate.AddDays(-1);
          DateTime startDay = generateDate.AddDays(-7);
          DateTime priorEndDay = generateDate.AddDays(-8);

          string endDateStr = endDay.ToString("yyyyMMdd");
          string startDateStr = startDay.ToString("yyyyMMdd");
          string priorEndDateStr = priorEndDay.ToString("yyyyMMdd");
          string priorGenerateDateStr = startDateStr;

          //TRICATE Team Weekly-demo.xls
          //excelHelper.SaveReportFile(saveFullPath,excelHelper.getFilenameByArg(-1, "Track Weekly(Thursday)") + @"\" + "TRICARE Team Weekly Report -" + startDate + "-" + endDate + ".xlsx");
          // TRICARE Weekly dump(thursdat)
          string packagePath = System.IO.Directory.GetCurrentDirectory().ToString().Replace(@"bin\Debug", @"SSISPackage\");//获取SSISPackage的目录。
          string ssisPathTricareThursday = packagePath + "tricateWeeklyThursday.dtsx";
          string excelFullNameTricareThursday = trackWeekly4Path + endDateStr + @"\" + "Track Wise Weekly Dump-" + startDateStr + "-" + endDateStr + ".xls";
          string excelTricarePaht = trackWeekly4Path + endDateStr;
          if (!Directory.Exists(excelTricarePaht))
          {
              Directory.CreateDirectory(excelTricarePaht);
          }
          ExcelHelper excelHelper = new ExcelHelper();
          excelHelper.ExportDataBySSIS(ssisPathTricareThursday, excelFullNameTricareThursday);

          ////handle tricare report sheet
          //string priorReportPath = trackWeekly4Path + priorEndDateStr + @"\Tricare Benchmark Report -" + startDateStr + ".xls";
          //string ReportPath = trackWeekly4Path + endDateStr + @"\Tricare Benchmark Report -" + generateDate.ToString("yyyyMMdd") + ".xls";
          //File.Copy(priorReportPath, ReportPath,true);
          //_ExcelApp = new Microsoft.Office.Interop.Excel.Application();
          //_ExcelWBook.CheckCompatibility = false;//Add for Diable Compatibility for saving excel
          //_ExcelWBook = (_Workbook)(_ExcelApp.Workbooks.Open(ReportPath, Missing.Value, Missing.Value
          //            , Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value,
          //            Missing.Value, Missing.Value, Missing.Value));
          //_Worksheet worksheet = _ExcelWBook.Worksheets["Dashboard"];
          //worksheet.get_Range("E3", "I58").Value = null;
          //worksheet.get_Range("D3", "D58").Copy(Type.Missing);
          //worksheet.get_Range("I3", "I3").PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteValues, Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone,
          //    Type.Missing , Type.Missing );
          //worksheet.get_Range("D3", "D58").Value = null;
          //_ExcelWBook.Save();
          //_ExcelWBook.Close();
          //_ExcelApp.Quit();

          System.Diagnostics.Process.Start("explorer.exe", trackWeekly4Path + endDateStr);
      }

      /// <summary>
      /// Prepare file for tricare Monday
      /// </summary>
      /// <param name="generateDate"></param>
      public void PrepareFilesForRADTricareMonday(DateTime generateDate)
      {
          DateTime endDay = generateDate.AddDays(-1);
          DateTime startDay = generateDate.AddDays(-9);

          string endDateStr = endDay.ToString("yyyyMMdd");
          string startDateStr = startDay.ToString("yyyyMMdd");
          string priorGenerateDateStr = startDateStr;

          //TRICATE Team Weekly-demo.xls
          //excelHelper.SaveReportFile(saveFullPath,excelHelper.getFilenameByArg(-1, "Track Weekly(Thursday)") + @"\" + "TRICARE Team Weekly Report -" + startDate + "-" + endDate + ".xlsx");
          // TRICARE Weekly dump(thursdat)
          string ssisPathTricareThursday = packagePath + "tricateWeeklyThursday.dtsx";
          string excelFullNameTricareThursday = trackWeekly1Path + endDateStr + @"\" + "Tricare Monday Weekly Dump-" + startDateStr + "-" + endDateStr + ".xls";
          string excelTricarePaht = trackWeekly1Path + endDateStr;
          if (!Directory.Exists(excelTricarePaht))
          {
              Directory.CreateDirectory(excelTricarePaht);
          }
          ExcelHelper excelHelper = new ExcelHelper();
          excelHelper.ExportDataBySSIS(ssisPathTricareThursday, excelFullNameTricareThursday);

          System.Diagnostics.Process.Start("explorer.exe", trackWeekly1Path + endDateStr);
      }

      /// <summary>
      /// Prepare file for tricare Monday
      /// </summary>
      /// <param name="generateDate"></param>
      public void PrepareFilesForCorpTricareMonday(DateTime generateDate)
      {
          DateTime endDay = generateDate.AddDays(-3);
          DateTime startDay = generateDate.AddDays(-9);

          string endDateStr = endDay.ToString("yyyyMMdd");
          string startDateStr = startDay.ToString("yyyyMMdd");
          string priorGenerateDateStr = startDateStr;

          string ssisPathCorpTricare = packagePath + "tricateCorpWeeklyMonday.dtsx";
          string excelCorpTricarePath = corpTrackWeekly1Path + startDay.ToString("MMdd") + "-" + endDay.ToString("MMdd");
          string excelFullNameCorpTricare = excelCorpTricarePath + @"\" + "Corp Tricare Weekly New dump " + startDateStr + "-" + endDateStr + ".xls";

          if (!Directory.Exists(excelCorpTricarePath))
          {
              Directory.CreateDirectory(excelCorpTricarePath);
          }
          ExcelHelper excelHelper = new ExcelHelper();
          excelHelper.ExportDataBySSIS(ssisPathCorpTricare, excelFullNameCorpTricare);

          System.Diagnostics.Process.Start("explorer.exe", excelCorpTricarePath);
      }

      /// <summary>
      /// handle tricare report sheet
      /// </summary>
      /// <param name="generateDate"></param>
      public string handleBenchmarkReportSheet(DateTime generateDate)
      {
          string errorMsg = string.Empty;
          DateTime endDay = generateDate.AddDays(-1);
          DateTime startDay = generateDate.AddDays(-7);
          DateTime priorEndDay = generateDate.AddDays(-8);
          try
          {

              string endDateStr = endDay.ToString("yyyyMMdd");
              string startDateStr = startDay.ToString("yyyyMMdd");
              string priorEndDateStr = priorEndDay.ToString("yyyyMMdd");
              string priorGenerateDateStr = startDateStr;

              string priorReportPath = trackWeekly4Path + priorEndDateStr + @"\Tricare Benchmark Report -" + startDateStr + ".xls";
              string ReportPath = trackWeekly4Path + endDateStr + @"\Tricare Benchmark Report -" + generateDate.ToString("yyyyMMdd") + ".xls";
              string reportTempPath = trackWeekly4Path + endDateStr + @"\TRICATE Team Weekly-" + startDateStr + "-" + endDateStr + ".xls";
              File.Copy(priorReportPath, ReportPath, true);
              _ExcelApp = new Microsoft.Office.Interop.Excel.Application();
              _ExcelWBook = (_Workbook)(_ExcelApp.Workbooks.Open(ReportPath, Missing.Value, Missing.Value
                          , Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                          Missing.Value, Missing.Value, Missing.Value));
              _ExcelWBook.CheckCompatibility = false;//Add for Diable Compatibility for saving excel
              _Worksheet worksheet = _ExcelWBook.Worksheets["Dashboard"];
              DateTime originalDate = Convert.ToDateTime("2013-1-3");
              int weekNum = (int)(generateDate - originalDate).Days / 7;
              string weeklyMsg = "Week-" + weekNum.ToString() + ": " + startDay.ToString("MM/dd/yyyy") + "-" + endDay.ToString("MM/dd/yyyy");
              worksheet.get_Range("D1", "D1").Value = weeklyMsg;
              worksheet.get_Range("E3", "I58").Value = null;
              worksheet.get_Range("D3", "D58").Copy(Type.Missing);
              worksheet.get_Range("I3", "I58").PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteValues, Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone,
                  Type.Missing, Type.Missing);
              worksheet.get_Range("D3", "D58").Value = null;

              _Workbook workbook2 = (_Workbook)(_ExcelApp.Workbooks.Open(reportTempPath, Missing.Value, Missing.Value
                          , Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                          Missing.Value, Missing.Value, Missing.Value));

              _Worksheet worksheet2 = workbook2.Worksheets["Benchmark Result"];
              worksheet2.get_Range("D3", "H58").Copy(Type.Missing);
              worksheet.get_Range("D3", "H58").PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteValues, Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone,
                 Type.Missing, Type.Missing);
              worksheet.Activate();
              worksheet.get_Range("A1", "A1").Select();

              _ExcelWBook.Save();
              _ExcelWBook.Close();
              workbook2.Close();
              _ExcelApp.Quit();
          }
          catch (Exception ex)
          {
              errorMsg = ex.Message;
          }

          return errorMsg;

      }




      public void ExportDataToExcel_New(string saveFileFullPath, string fileName, string[] tableName, string[] sheetName, DataSet ds)
      {
          PreExitExcel();
          string filename = fileName.Substring(0, fileName.LastIndexOf("."));
          try
          {
              _ExcelApp = new Microsoft.Office.Interop.Excel.Application();
              _ExcelWBook = (_Workbook)(_ExcelApp.Workbooks.Open(excelTemplatePath + filename, Missing.Value, Missing.Value
                      , Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                      Missing.Value, Missing.Value, Missing.Value));

              for (int i = 0; i < tableName.Length; i++)
              {
                  _IniRow = 3;
                  _sheetName = sheetName[i].ToString();
                  _ExcelWSheet = (_Worksheet)_ExcelWBook.Sheets[sheetName[i].ToString()];
                  int rowcount = ds.Tables[i].Rows.Count;
                  for (int j = 0; j < ds.Tables[i].Rows.Count; j++)
                  {
                      for (int k = 0; k < ds.Tables[i].Columns.Count; k++)
                      {
                          _ExcelWSheet.Cells[_IniRow, k + 1] = ds.Tables[i].Rows[j][k].ToString();
                      }
                      _IniRow++;
                  }
              }

              ExcelHelper excelHelper = new ExcelHelper();
              _ExcelApp.Rows.RowHeight = "15";

              //ExApp.DisplayAlerts = false;
              _ExcelWBook.CheckCompatibility = false;//Add for Diable Compatibility for saving excel
              _ExcelWBook.SaveAs(saveFileFullPath, XlFileFormat.xlWorkbookNormal,
                  null, null, false, false, XlSaveAsAccessMode.xlExclusive, false, false, null, null, null);
              _ExcelWBook.Close(Missing.Value, Missing.Value, Missing.Value);
              _ExcelApp.Workbooks.Close();
              _ExcelApp.Quit();

              Marshal.ReleaseComObject(_ExcelWSheet);
              Marshal.ReleaseComObject(_ExcelWBook);
              Marshal.ReleaseComObject(_ExcelApp);
          }
          catch (Exception e)
          {
              _ExcelWBook.Close(Missing.Value, Missing.Value, Missing.Value);
              _ExcelApp.Workbooks.Close();
              _ExcelApp.Quit();
              throw e;
          }
      }
    }
}
