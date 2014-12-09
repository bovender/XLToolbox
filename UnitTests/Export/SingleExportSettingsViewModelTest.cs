using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NUnit.Framework;
using Bovender.Mvvm.Messaging;
using XLToolbox.Export;
using XLToolbox.Excel.Instance;

namespace XLToolbox.UnitTests.Export
{
    [TestFixture]
    class SingleExportSettingsViewModelTest
    {
        SingleExportSettingsViewModel svm;

        [SetUp]
        void SetUp()
        {
            svm = new SingleExportSettingsViewModel();
        }

        [Test]
        public void PreserveAspectWidth()
        {
            svm.Height = 100;
            svm.Width = 200;
            svm.PreserveAspect = true;
            svm.Height = 200;
            Assert.AreEqual(svm.Height * 2, svm.Width,
                "Width was not correctly changed.");
        }

        [Test]
        public void PreserveAspectHeight()
        {
            svm.Height = 100;
            svm.Width = 200;
            svm.PreserveAspect = true;
            svm.Width = 400;
            Assert.AreEqual(svm.Width / 2, svm.Height,
                "Height was not correctly changed.");
        }
        
        [Test]
        public void ExportCommandDisabledWithoutSelection()
        {
            Assert.IsFalse(svm.ExportCommand.CanExecute(null),
                "Export command should be disabled if there is no selection.");
            using (new ExcelInstance())
            {
                ExcelInstance.CreateWorkbook();
                Assert.IsTrue(svm.ExportCommand.CanExecute(null),
                    "Export command should be enabled if something is selected.");
            }
        }

        [Test]
        public void ExportCommand()
        {
            using (new ExcelInstance())
            {
                ExcelInstance.CreateWorkbook();
                bool fnMsgSent = false;
                svm.ChooseFileNameMessage.Sent +=
                    (object sender, MessageArgs<StringMessageContent> args) =>
                    {
                        fnMsgSent = true;
                        // args.Respond();
                    };
                svm.ExportCommand.Execute(null);
                Assert.IsTrue(fnMsgSent, "ChooseFileNameMessage was not sent.");
            }
        }
    }
}
