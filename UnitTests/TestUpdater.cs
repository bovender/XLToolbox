using System;
using XLToolbox.Version;
using NUnit.Framework;

namespace XLToolbox.Test
{
    [TestFixture]
    public class TestUpdater
    {
        [Test]
        public void FetchVersionInformation()
        {
            Updater u = new Updater();
            u.FetchVersionInformation();
        }
    }
}
