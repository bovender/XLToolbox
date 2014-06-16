using System;
using NUnit.Framework;
using Microsoft.Office.Interop.Excel;

namespace XLToolbox.Test
{
    /// <summary>
    /// Provides an NUnit test context with two static members to allow easy access
    /// to an Excel application and a standard workbook in it.
    /// </summary>
    [SetUpFixture]
    public class TestContext
    {
        public static Application Excel { get; private set; }
        public static Workbook Workbook { get; private set; }

        public TestContext() {}

        [SetUp]
        public void Initialize()
        {
            Excel = new Application();
            Excel.Visible = Properties.Settings.Default.RunVisibly;
            Workbook = Excel.Workbooks.Add();
        }

        [TearDown]
        public void Teardown()
        {
            Excel.DisplayAlerts = false;
            Excel.Quit();
        }
    }
}
