using System;
using NUnit.Framework;
using XLToolbox.WorkbookStorage;
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
                Store storage = new Store(excel);
            });
        }

        [Test]
        [ExpectedException(typeof(WorkbookStorageException), ExpectedMessage = "Invalid storage context")]
        public void InvalidContextCausesException()
        {
            Store storage = new Store(TestContext.Workbook);
            storage.Context = "made-up context that does not exist";
        }

        [Test]
        public void StoreAndRetrieveInteger()
        {
            Store storage1 = new Store(TestContext.Workbook);
            Store storage2 = new Store(TestContext.Workbook);
            string key = "Some property key";
            int i = 123;
            storage1.UseActiveSheet();
            storage1.Put(key, i);
            storage2.Context = storage1.Context;
            int j = storage2.Get(key);
            Assert.AreEqual(i, j);
        }
    }
}
