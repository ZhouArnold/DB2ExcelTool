using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Windows;
using Microsoft.Office.Interop.Excel;
using DtsRunTime = Microsoft.SqlServer.Dts.Runtime.Wrapper;
using Microsoft.SqlServer.Dts.Runtime.Wrapper;

namespace WindowsFormsApplication3
{
   public  class ExcelHelper
    {
        public static string packageErrorMsg = string.Empty;
        public string getFilenamePF(DateTime generateDate)
        {
            DateTime targetDate = generateDate.AddDays(-1);
            string targetDateStr = targetDate.ToString("yyyyMMdd");
            string filename = string.Format(@"D:\PSTWork\Reports\{0}", targetDateStr);
            if (!Directory.Exists(filename))
            {
                Directory.CreateDirectory(filename);
            }
            return filename;
        }

        public string getFilename(DateTime generateDate)
        {
            DateTime targetDate = generateDate.AddDays(-1);
            string targetDateStr = targetDate.ToString("yyyyMMdd");
            string fileNamePath = string.Format(@"D:\PSTWork\AutoReports\Final reports\{0}", targetDateStr);
            if (!Directory.Exists(fileNamePath))
            {
                Directory.CreateDirectory(fileNamePath);
            }
            return fileNamePath;
        }

        /// <summary>
        /// Exports data by ssis.
        /// </summary>
        /// <param name="ssisPath">package full name(including path)</param>
        /// <param name="excelFullName">excel full name(including path)</param>
        public void ExportDataBySSIS(string ssisPath, string excelFullName)
        {
            #region Old logic
            //Microsoft.SqlServer.Dts.Runtime.Application app = new Microsoft.SqlServer.Dts.Runtime.Application();
            //Package package = app.LoadPackage(ssisPath, null);
            //string excelDest = string.Format("Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties=\"EXCEL 8.0;HDR=YES\";", excelFullName);
            //Microsoft.SqlServer.Dts.Runtime.Connections conns = package.Connections;
            ////package.
            //package.Connections["DestinationConnectionExcel"].ConnectionString = excelDest;
            //DTSExecResult r = package.Execute();
            #endregion
            //Add an local DLL
            //DLL reference path: C:\Program Files\Microsoft SQL Server\100\SDK\Assemblies
            //DLL Name: Microsoft.SQLServer.DTSRuntimeWrap.dll        

            //packageErrorMsg = string.Empty;

            DtsRunTime.Application app = new DtsRunTime.Application();
            DtsRunTime.IDTSPackage100 package = app.LoadPackage(ssisPath, false, null);
            string excelDest = string.Format("Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties=\"EXCEL 8.0;HDR=YES\";", excelFullName);
            package.Connections["DestinationConnectionExcel"].ConnectionString = excelDest;
            PackageEvenet pevent = new PackageEvenet();
            DtsRunTime.DTSExecResult result = package.Execute(null, null, pevent, null, null);

            if (DtsRunTime.DTSExecResult.DTSER_FAILURE == result)
            {
                //MessageBox.Show("SSIS Run Failed and error message is " + packageErrorMsg);
            }
        }

        class PackageEvenet : DtsRunTime.IDTSEvents100
        {
            public void OnError(IDTSRuntimeObject100 pSource,
                int ErrorCode,
                string SubComponent,
                string Description,
                string HelpFile,
                int HelpContext,
                string IDOfInterfaceWithError,
                out bool pbCancel)
            {
                pbCancel = false;
                packageErrorMsg = packageErrorMsg + Description;
            }

            public void OnBreakpointHit(IDTSBreakpointSite100 pBreakpointSite, IDTSBreakpointTarget100 pBreakpointTarget)
            {
            }

            public void OnCustomEvent(IDTSTaskHost100 pTaskHost, string EventName, string EventText, ref object[] ppsaArguments, string SubComponent, ref bool pbFireAgain) { }

            public void OnExecutionStatusChanged(IDTSExecutable100 pExec, DTSExecStatus newStatus, ref bool pbFireAgain) { }

            public void OnInformation(IDTSRuntimeObject100 pSource, int InformationCode, string SubComponent, string Description, string HelpFile, int HelpContext, string IDOfInterfaceWithError, ref bool pbFireAgain) { }

            public void OnPostExecute(IDTSExecutable100 pExec, ref bool pbFireAgain) { }

            public void OnPostValidate(IDTSExecutable100 pExec, ref bool pbFireAgain) { }

            public void OnPreExecute(IDTSExecutable100 pExec, ref bool pbFireAgain) { }

            public void OnPreValidate(IDTSExecutable100 pExec, ref bool pbFireAgain) { }

            public void OnProgress(IDTSTaskHost100 pTaskHost, string ProgressDescription, int PercentComplete, int ProgressCountLow, int ProgressCountHigh, string SubComponent, ref bool pbFireAgain) { }

            public void OnQueryCancel(out bool pbCancel) { pbCancel = false; }

            public void OnTaskFailed(IDTSTaskHost100 pTaskHost) { }

            public void OnVariableValueChanged(IDTSContainer100 pContainer, IDTSVariable100 pVariable, ref bool pbFireAgain) { }

            public void OnWarning(IDTSRuntimeObject100 pSource, int WarningCode, string SubComponent, string Description, string HelpFile, int HelpContext, string IDOfInterfaceWithError)
            { }

        }

    }
}
