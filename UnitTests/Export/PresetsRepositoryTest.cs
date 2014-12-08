using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using XLToolbox.Export;
using NUnit.Framework;

namespace XLToolbox.UnitTests.Export
{
    [TestFixture]
    public class PresetsRepositoryTest
    {
        [Test]
        public void StoreAndRetrieve()
        {
            string testName = "test settings";
            using (PresetsRepository repository = new PresetsRepository())
            {
                Preset settings = new Preset() { Dpi = 300, ColorSpace = ColorSpace.Cmyk, Name = testName };
                repository.Add(settings);
            }
            using (PresetsRepository repository = new PresetsRepository())
            {
                Preset settings = repository.ExportSettings[0];
                Assert.AreEqual(testName, settings.Name,
                    "Retrieved export settings have different name than previously stored.");
            }
        }
    }
}
