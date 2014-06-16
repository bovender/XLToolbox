using System;
using NUnit.Framework;
using XLToolbox.Core;
using Microsoft.Office.Interop.Excel;

namespace XLToolbox.Test
{
    [TestFixture]
    public class TestWorkbookStorage
    {
        [Test]
        public void AppWithoutWorkbookDoesNotCrashStorage()
        {
            Assert.DoesNotThrow(delegate()
            {
                Application excel = new Application();
                WorkbookStorage storage = new WorkbookStorage(excel);
            });
        }

        [Test]
        [ExpectedException(typeof(WorkbookStorageException), ExpectedMessage = "Invalid storage context")]
        public void InvalidContextCausesException()
        {
            WorkbookStorage storage = new WorkbookStorage(TestContext.Workbook);
            storage.Context = "made-up context that does not exist";
        }

        [Test]
        public void StoreAndRetrieveInteger()
        {
            WorkbookStorage storage = new WorkbookStorage(TestContext.Workbook);
            string key = "Some property key";
            int i = 123;
            storage.UseActiveSheet();
            storage.Store(key, i);
            int j = storage.Retrieve(key);
            Assert.AreEqual(i, j);
        }
    }
}
