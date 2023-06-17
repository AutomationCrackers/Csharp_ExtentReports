using AventStack.ExtentReports;
using AventStack.ExtentReports.Reporter;
using OpenQA.Selenium;
using System;
using System.IO;

namespace Csharp_ExtentReports
{
    public class ExtentHTMLReportGenerator
    {
        private static string ExtentReportpath, report;
        private static ExtentReports extent;
        public static IWebDriver driver;
        public static ExtentTest test;

        //Step 1 Define Extent Report Path
        public static void GetExtentReportPath()
        {
            string currentDirectoryPath = Environment.CurrentDirectory;
            string actualPath = currentDirectoryPath.Substring(0, currentDirectoryPath.LastIndexOf("bin"));
            string projectPath = new Uri(actualPath).LocalPath;
            ExtentReportpath = projectPath + "\\ExtentReports\\";
        }

        //Step 2 - Delete the Old Exisitng Reports inside the Path
        public static void DeleteOldExtentReports()
        {
            GetExtentReportPath();
            if (Directory.Exists(ExtentReportpath))
            {
                string[] ExtentReports = Directory.GetFiles(ExtentReportpath);
                foreach (string Reports in ExtentReports)
                {
                    File.Delete(ExtentReportpath);
                }
            }
        }
        //Step 3 - Define the Extent Report path where it has to be created, font, name, styling
        public static void ConfigureExtentReport()
        {
            DeleteOldExtentReports();
            GetExtentReportPath();
            report = ExtentReportpath + "\\ExtentReports\\Report.html";
            var htmlreport = new ExtentHtmlReporter(report);

            //Configue all the Report Properties
            htmlreport.Config.ReportName = "Automation Report";
            htmlreport.Config.DocumentTitle = "Automation Test Report";
            htmlreport.Config.Theme = AventStack.ExtentReports.Reporter.Configuration.Theme.Standard;
            extent = new AventStack.ExtentReports.ExtentReports();
            extent.AttachReporter(htmlreport);
        }

        //Step 4 - Flush Extent Report
        public static void CloseExtentReport()
        {
            extent.Flush();
        }

        //Step 5 - Rename Index.Html and customize own name
        public static void RenameExtentReport()
        {
            GetExtentReportPath();
            System.IO.File.Move(ExtentReportpath + @"\index.html", ExtentReportpath + "Automation Test Report" + DateTime.Now.ToString("d-M-yyyy")
                + "Time" + DateTime.Now.ToString("hhmmss tt") + ".html");
            //Note: you can change "Automation Test Report" name to your customized name or Feature Speicifc or scenario specific as well
        }
        
        //Capture Screenshot using Media Entity
        public static MediaEntityModelProvider CaptureScreenshot() 
        {
            var screenshot = ((ITakesScreenshot)driver).GetScreenshot().AsBase64EncodedString;
            return MediaEntityBuilder.CreateScreenCaptureFromBase64String(screenshot).Build();
        }

        // Step 6 - Attach Screenshot to Extent Report
        //Call this Method whereever you want to take screenshot
        public static void LogScreeshotInReport()
        {
            MediaEntityModelProvider ss = CaptureScreenshot();
            test.Log((Status)5, ss);
        }
    }
}
