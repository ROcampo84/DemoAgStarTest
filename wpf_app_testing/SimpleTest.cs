using System;
using System.Diagnostics;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium;
using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Windows;
using OpenQA.Selenium.Support.UI;

namespace agstar_testing
{
    [TestClass]
    public class SimpleTest
    {
        static WindowsDriver<WindowsElement> sessionWpfApp;

        static TestContext objTestContext;

        protected const string winAppDriverURL = "http://127.0.0.1:4723";

        private const string AppId = @"C:\Program Files (x86)\AgStar\AgStar.exe";

        [ClassInitialize]
        public static void PrepareForTestingAlarms(TestContext testContext)
        {
            AppiumOptions capCalc = new AppiumOptions();

            capCalc.AddAdditionalCapability("app", AppId);

            sessionWpfApp = new WindowsDriver<WindowsElement>(new Uri(winAppDriverURL), capCalc);

            objTestContext = testContext;
        }

        [ClassCleanup]
        public static void CleanupAfterAllAlarmsTests()
        {
            if (sessionWpfApp != null)
            {
                //sessionWpfApp.Quit();
            }
        }

        [TestMethod,
            DataSource("System.Data.OleDb",
                       @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\DOCUMENTOS\Cursos\Appium_WinAppDriver\AgStarTesting\agstar_automation_testing.xlsx;Extended Properties=""Excel 12.0 Xml;HDR = YES"";",
                       "SerialNumbers$",
                       DataAccessMethod.Sequential)]
        public void AgStarSerialTest()
        {
            sessionWpfApp.FindElementByAccessibilityId("dlgSetLicense").Click(); //Focus on Textbox to type serial number

            WebDriverWait waitForMe = new WebDriverWait(sessionWpfApp, TimeSpan.FromSeconds(10));

            var txtPSetLicense = sessionWpfApp.FindElementByAccessibilityId("dlgSetLicense");

            waitForMe.Until(pred => txtPSetLicense.Displayed);

            String serialNumber = Convert.ToString(objTestContext.DataRow["Data Input"]);

            Debug.WriteLine($"Input from Excel: {serialNumber}");

            txtPSetLicense.SendKeys(serialNumber);

            txtPSetLicense.SendKeys(Keys.Enter);

            try
            {
                sessionWpfApp.FindElementByAccessibilityId("2").Click(); //Button OK from Error Message Box
            }
            catch (Exception)
            {
                //In case that element doesn't exists.
            }

            Thread.Sleep(2000);
        }

        /*
        [TestMethod,
            DataSource("System.Data.OleDb",
                       @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\DOCUMENTOS\Cursos\Appium_WinAppDriver\AgStarTesting\agstar_automation_testing.xlsx;Extended Properties=""Excel 12.0 Xml;HDR = YES"";",
                       "DbNames$",
                       DataAccessMethod.Sequential)]
        public void AgStarDbNameTest()
        {
            sessionWpfApp.FindElementByAccessibilityId("dbName").Click(); //ComboBox to choose database name

            WebDriverWait waitForMe = new WebDriverWait(sessionWpfApp, TimeSpan.FromSeconds(10));

            var txtDbName = sessionWpfApp.FindElementByAccessibilityId("dbName");

            waitForMe.Until(pred => txtDbName.Displayed);

            String dbName = Convert.ToString(objTestContext.DataRow["Data Input"]);

            Debug.WriteLine($"Input from Excel: {dbName}");

            txtDbName.SendKeys(dbName);

            txtDbName.SendKeys(Keys.Enter);

            Thread.Sleep(5000);
        }
        */
    }
}
