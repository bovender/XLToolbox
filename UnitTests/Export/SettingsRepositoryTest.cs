using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using XLToolbox.Export;
using NUnit.Framework;

namespace XLToolbox.UnitTests.Export
{
    [TestFixture]
    public class SettingsRepositoryTest
    {
        [Test]
        public void StoreAndRetrieve()
        {
            string testName = "test settings";
            using (SettingsRepository repository = new SettingsRepository())
            {
                Settings settings = new Settings() { Dpi = 300, ColorSpace = ColorSpace.Cmyk, Name = testName };
                repository.Add(settings);
            }
            using (SettingsRepository repository = new SettingsRepository())
            {
                Settings settings = repository.ExportSettings[0];
                Assert.AreEqual(testName, settings.Name,
                    "Retrieved export settings have different name than previously stored.");
            }
        }
    }
}
