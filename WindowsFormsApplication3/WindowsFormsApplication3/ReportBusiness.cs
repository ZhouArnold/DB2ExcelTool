using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;

namespace WindowsFormsApplication3
{
    class ReportBusiness
    {
        private static string packagePath = System.IO.Directory.GetCurrentDirectory().ToString().Replace(@"bin\Debug", @"SSISPackage\");//获取SSISPackage的目录。
        private static string excelTemplatePath = System.IO.Directory.GetCurrentDirectory().ToString().Replace(@"bin\Debug", @"ExcelTemplate\");//获取Exceltemplate的目录。
        private static string finalReportsPath = "D:\\PSTWork\\AutoReports\\Final reports\\";
        private static string trackWeekly4Path = "D:\\PSTWork\\Track Weekly(Thursday)\\";
        private static string trackWeekly1Path = "D:\\PSTWork\\Track Weekly(Monday)\\";
        private static string corpTrackWeekly1Path = "D:\\PSTWork\\Track Corp Weekly(Monday)\\";

        /// <summary>
        /// Generates daily reports for multiple days.
        /// </summary>
        /// <param name="generateDate"></param>
        /// <param name="ischecked"></param>
        public void GenerateDailyRptForMultiDays(DateTime generateDate)
        {
            DateTime todayDate = DateTime.Now;
            int delayDays = (generateDate - todayDate).Days;//get number for store procedure parameter
            int startDelayDays = delayDays - 2;//get number for store procedure parameter
            string StartDateStr = generateDate.AddDays(-3).ToString("yyyy-MM-dd");
            string EndDateStr = generateDate.AddDays(-1).ToString("yyyy-MM-dd");
            string filename1W = "Service and TricareDaily.xls";
            string[] sheetnameW = new string[5] { "ServiceMgRpt", "AssignedInAndOut_Service", "IM", "SR", "PM" };
            string[] ProcNamesW = new string[5] { 
                                                    "ServiceMgRptByNumbers"
                                                    , "sp_TaskInAndOut_ServiceRPT"
                                                    , "IMRpt_TricareForMulti"
                                                    , "SRRpt_TricareForMulti"
                                                    , "PBLMRpt_TricareForMulti" };
            string[,] ParamsW = new string[5, 2] { 
                                                    { delayDays.ToString(), startDelayDays.ToString() }
                                                    , { "", "" }
                                                    , { delayDays.ToString(), startDelayDays.ToString() }
                                                    , { delayDays.ToString(), startDelayDays.ToString() }
                                                    , { delayDays.ToString(), startDelayDays.ToString() } 
                                                };
            //prds.ExecuteProcWithParam(filename1W, ProcNamesW, sheetnameW, ParamsW, ischecked);
            DBAccess dbAccess = new DBAccess();
            dbAccess.ExecuteProcWithParam(filename1W, ProcNamesW, sheetnameW, ParamsW, true, generateDate);

            //Annuity and Compensation daily reports
            string AEfilename = "Annuity and EBusiness.xls";
            string[] AEWorksheetName = new string[4] { 
                                                        "AnnuityRpt",
                                                        "AnnuityStatus" ,
                                                        "AnnuityFinal",
                                                        "EBusinessRpt Grouped"
                                                    };
            //Porcedure names and datatable names
            string[] AEProcName = new string[6] {                 
                                                    "AnnuityRpt",         //{1,1}        
                                                    "AnnuityOpenPeriod",   //{51,2}              
                                                    "AnnuityUrgent",  //{61,2}
                                                    "AnnuityStatus", //{1,1}
                                                    "AnnuityStatusUrgent" ,//{73,2}
                                                    "EBusinessRpt"
                                                }; //{123,2}
            //Params of procedures
            string[,] AEParam = new string[6, 3] { 
                                                    { StartDateStr,EndDateStr , "1" },
                                                    { StartDateStr,EndDateStr, "1" },
                                                    { StartDateStr,EndDateStr, "1" },
                                                    {"","",""},
                                                    {"","",""},
                                                    { StartDateStr,EndDateStr, "1" }
                                                 };
            //Row number of excel when write the ds data into excel
            int[] sheetWriteStartRow = new int[6] { 1, 89, 99, 1, 109, 2 };

            DataSet dsOfAE = dbAccess.getDSByProcs(AEProcName, AEParam);
            ExcelBus excelBiz = new ExcelBus();
            excelBiz.AEExportDataToExcel(AEfilename, AEProcName, AEWorksheetName, dsOfAE, sheetWriteStartRow, generateDate);

            //EBusiness backlog tickets
            string[] ProcName = new string[1] { "getEBusinessData" };
            string[,] Param = new string[1, 1] { { "" } };
            string[] sheetnameE = new string[1] { "EBusiness Raw Data" };
            string filenameE = "EBusiness ticket.xls";
            dbAccess.ExecuteProcWithParam(filenameE, ProcName, sheetnameE, Param, false, generateDate);

            //Annuity backlog tickets
            string[] ProcNameAnnuity = new string[1] { "getAnnuityData" };
            string[,] ParamAnnuity = new string[1, 1] { { "" } };
            string[] sheetnameAnnuity = new string[1] { "Annuity Raw Data" };
            string filenameAnnuity = "Annuity ticket.xls";
            dbAccess.ExecuteProcWithParam(filenameAnnuity, ProcNameAnnuity, sheetnameAnnuity, ParamAnnuity, false, generateDate);
        }



        public void GenerateDailyRptForSingleDays(DateTime generateDate)
        {
            DateTime todayDate = DateTime.Now;
            int delayDays = (generateDate - todayDate).Days;//get number for store procedure parameter
            string endDateStr = generateDate.AddDays(-1).ToString("yyyy-MM-dd");
            string filename1W = "Service and TricareDaily.xls";
            string[] sheetnameW = new string[5] { "ServiceMgRpt", "AssignedInAndOut_Service", "IM", "SR", "PM" };
            string[] ProcNamesW = new string[5] {  
                                                    "ServiceMgRptByNumber"
                                                    , "sp_TaskInAndOut_ServiceRPT"
                                                    , "IMRpt_Tricare"
                                                    , "SRRpt_Tricare"
                                                    , "PBLMRpt_Tricare" 
                                                 };
            string[,] ParamsW = new string[5, 1] { 
                                                    { delayDays.ToString() }
                                                    , { "" }
                                                    , { delayDays.ToString() }
                                                    , { delayDays.ToString() }
                                                    , { delayDays.ToString() } 
                                                };
            //prds.ExecuteProcWithParam(filename1W, ProcNamesW, sheetnameW, ParamsW, ischecked);
            DBAccess dbAccess = new DBAccess();
            dbAccess.ExecuteProcWithParam(filename1W, ProcNamesW, sheetnameW, ParamsW, false, generateDate);

            //Annuity and Compensation daily reports
            string AEfilename = "Annuity and EBusiness.xls";
            string[] AEWorksheetName = new string[4] { 
                                                        "AnnuityRpt",
                                                        "AnnuityStatus" ,
                                                        "AnnuityFinal",
                                                        "EBusinessRpt Grouped"
                                                    };
            //Porcedure names and datatable names
            string[] AEProcName = new string[6] {                 
                                                    "AnnuityRpt",         //{1,1}        
                                                    "AnnuityOpenPeriod",   //{51,2}              
                                                    "AnnuityUrgent",  //{61,2}
                                                    "AnnuityStatus", //{1,1}
                                                    "AnnuityStatusUrgent" ,//{73,2}
                                                    "EBusinessRpt"
                                                }; //{123,2}
            //Params of procedures
            string[,] AEParam = new string[6, 3] { 
                                                    { "1900/01/01",endDateStr , "0" },
                                                    { "1900/01/01",endDateStr, "0" },
                                                    { "1900/01/01",endDateStr, "0" },
                                                    {"","",""},
                                                    {"","",""},
                                                    { "1900/01/01",endDateStr, "0" }
                                                 };
            //Row number of excel when write the ds data into excel
            int[] sheetWriteStartRow = new int[6] { 1, 89, 99, 1, 109, 2 };

            DataSet dsOfAE = dbAccess.getDSByProcs(AEProcName, AEParam);
            ExcelBus excelBiz = new ExcelBus();
            excelBiz.AEExportDataToExcel(AEfilename, AEProcName, AEWorksheetName, dsOfAE, sheetWriteStartRow, generateDate);

            //EBusiness backlog tickets
            string[] ProcName = new string[1] { "getEBusinessData" };
            string[,] Param = new string[1, 1] { { "" } };
            string[] sheetnameE = new string[1] { "EBusiness Raw Data" };
            string filenameE = "EBusiness ticket.xls";
            dbAccess.ExecuteProcWithParam(filenameE, ProcName, sheetnameE, Param, false, generateDate);

            //Annuity backlog tickets
            string[] ProcNameAnnuity = new string[1] { "getAnnuityData" };
            string[,] ParamAnnuity = new string[1, 1] { { "" } };
            string[] sheetnameAnnuity = new string[1] { "Annuity Raw Data" };
            string filenameAnnuity = "Annuity ticket.xls";
            dbAccess.ExecuteProcWithParam(filenameAnnuity, ProcNameAnnuity, sheetnameAnnuity, ParamAnnuity, false, generateDate);
        }

        public void GenerateServiceRemedyGroupLevel(DateTime generateDate)
        {
            DBAccess dbAccess = new DBAccess();
            //Generate service remedy group level report for Georg
            string[] ServiceRemedyGourpProc = new string[1] { "ServiceGroupLevel" };
            string[,] ServiceRemedyGourpParam = new string[1, 1] { { "" } };
            string[] ServiceRemedyGourpRptTab = new string[1] { "Report" };
            string ServiceRemedyGourpRptSheet = "Service Remedy Group Level Report.xls";
            dbAccess.ExecuteProcWithParam(ServiceRemedyGourpRptSheet, ServiceRemedyGourpProc, ServiceRemedyGourpRptTab, ServiceRemedyGourpParam, false, generateDate);

        }

        /// <summary>
        /// Generates tricarte weekly report on Thursday.
        /// </summary>
        public void GenerateTricareWeeklyRpt(DateTime generateDate, out string errorMsg)
        {
            DateTime todayDate = DateTime.Now;
            int delayDays = (generateDate - todayDate).Days;//get number for store procedure parameter
            int startDelayDays = delayDays - 6;//get number for store procedure parameter
            string[] ProcNamesW = new string[4] { "IMRpt_TricareForMulti", "SRRpt_TricareForMulti", "PBLMRpt_TricareForMulti", "TriBenchMarkRpt_MatchedBacklog" };
            string[,] ParamsW = new string[3, 2] { 
                                                    { delayDays.ToString(), startDelayDays.ToString() }
                                                    ,{ delayDays.ToString(), startDelayDays.ToString() }                                                    
                                                    ,{ delayDays.ToString(), startDelayDays.ToString() }
                                                 };//benchmark para add in ExecuteProcWithParam method
            string[] sheetnameW = new string[4] { "IM", "SR", "PM", "BenchMarkData" };
            string filename1W = "TRICATE Team Weekly-demo.xls";

            DBAccess dbAccess = new DBAccess();
            dbAccess.ExecuteProcWithParam(filename1W, ProcNamesW, sheetnameW, ParamsW, true, generateDate);

            DateTime EndDay = generateDate.AddDays(-1);
            DateTime StartDay = generateDate.AddDays(-7);

            string endDate = EndDay.ToString("yyyyMMdd");
            string startDate = StartDay.ToString("yyyyMMdd");
            string endDateShort = EndDay.ToString("MMdd");

            //TRICATE Team Weekly-demo.xls
            ExcelHelper excelHelper = new ExcelHelper();
            string sourceFilePath = finalReportsPath + endDate + @"\" + "TRICATE Team Weekly-demo-" + endDateShort + ".xls";
            string copyToPath = trackWeekly4Path + endDate;
            string copyToFullPath = copyToPath + @"\" + "TRICATE Team Weekly-" + startDate + "-" + endDate + ".xls";

            if (!Directory.Exists(copyToPath))
            {
                Directory.CreateDirectory(copyToPath);
            }
            File.Copy(sourceFilePath, copyToFullPath, true);

            ExcelBus excelBiz = new ExcelBus();
            errorMsg = excelBiz.handleBenchmarkReportSheet(generateDate);
        }


        /// <summary>
        /// Generates tricarte weekly report on Monday.
        /// </summary>
        public void GenerateTricareMonday(DateTime generateDate)
        {
            DateTime todayDate = DateTime.Now;
            int delayDays = (generateDate - todayDate).Days;//get number for store procedure parameter
            int startDelayDays = delayDays - 8;//get number for store procedure parameter
            string[] ProcNamesW = new string[3] { "IMRpt_TricareForMulti", "SRRpt_TricareForMulti", "PBLMRpt_TricareForMulti" };
            string[,] ParamsW = new string[3, 2] { 
                                                    { delayDays.ToString(), startDelayDays.ToString() }
                                                    ,{ delayDays.ToString(), startDelayDays.ToString() }                                                    
                                                    ,{ delayDays.ToString(), startDelayDays.ToString() }
                                                 };//benchmark para add in ExecuteProcWithParam method
            string[] sheetnameW = new string[3] { "IM", "SR", "PM" };
            string filename1W = "TRICATE Team Weekly-demo.xls";

            DBAccess dbAccess = new DBAccess();
            dbAccess.ExecuteProcWithParam(filename1W, ProcNamesW, sheetnameW, ParamsW, true, generateDate);

            DateTime EndDay = generateDate.AddDays(-1);
            DateTime StartDay = generateDate.AddDays(-9);

            string endDate = EndDay.ToString("yyyyMMdd");
            string startDate = StartDay.ToString("yyyyMMdd");
            string endDateShort = EndDay.ToString("MMdd");

            //TRICATE Team Weekly-demo.xls
            ExcelHelper excelHelper = new ExcelHelper();
            string sourceFilePath = finalReportsPath + endDate + @"\" + "TRICATE Team Weekly-demo-" + endDateShort + ".xls";
            string copyToPath = trackWeekly1Path + endDate;
            string copyToFullPath = copyToPath + @"\" + "TRICATE Team Weekly-" + startDate + "-" + endDate + ".xls";

            if (!Directory.Exists(copyToPath))
            {
                Directory.CreateDirectory(copyToPath);
            }
            File.Copy(sourceFilePath, copyToFullPath, true);
        }

        public void GenerateCorpTriMonday(DateTime generateDate)
        {
            DateTime ReprtStartDate = generateDate.AddDays(-9);
            DateTime ReprtEndDate = generateDate.AddDays(-3);

            string[] ProcNamesW = new string[5] { "Corp_IM", "Corp_SR", "Corp_PBI", "Corp_PKE", "Corp_WO" };
            string[,] ParamsW = new string[5, 2] { 
                                                    { ReprtStartDate.ToString("yyyy-MM-dd"), ReprtEndDate.ToString("yyyy-MM-dd") }
                                                    ,{ ReprtStartDate.ToString("yyyy-MM-dd"), ReprtEndDate.ToString("yyyy-MM-dd") }
                                                    ,{ ReprtStartDate.ToString("yyyy-MM-dd"), ReprtEndDate.ToString("yyyy-MM-dd") }
                                                    ,{ ReprtStartDate.ToString("yyyy-MM-dd"), ReprtEndDate.ToString("yyyy-MM-dd") }
                                                    ,{ ReprtStartDate.ToString("yyyy-MM-dd"), ReprtEndDate.ToString("yyyy-MM-dd") }
                                                 };//benchmark para add in ExecuteProcWithParam method
            string[] sheetnameW = new string[5] { "IM", "SR", "PBI", "PKE", "WO" };
            string templateReportfile = "Corp Tricare Weekly New -demo.xls";

            DBAccess dbAccess = new DBAccess();
            DataSet CorpDs = dbAccess.ExecuteProcWithParam_Navy(templateReportfile, ProcNamesW, sheetnameW, ParamsW, ReprtEndDate);

            string reportStartStr = ReprtStartDate.ToString("MMdd");
            string reportEndStr = ReprtEndDate.ToString("MMdd");
            ExcelBus excelBiz = new ExcelBus();
            string savePath = corpTrackWeekly1Path + reportStartStr + "-" + reportEndStr;
            string corpReportSavePath = savePath + @"\Corp Tricare Weekly New " + reportStartStr + "-" + reportEndStr + ".xls";

            if (!Directory.Exists(savePath))
            {
                Directory.CreateDirectory(savePath);
            }
            excelBiz.ExportDataToExcel_New(corpReportSavePath, templateReportfile, ProcNamesW, sheetnameW, CorpDs);
        }

    }
}
