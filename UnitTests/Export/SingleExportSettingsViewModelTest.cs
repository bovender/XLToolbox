using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NUnit.Framework;
using Bovender.Mvvm.Messaging;
using XLToolbox.Export;
using XLToolbox.Excel.Instance;
using XLToolbox.Export.ViewModels;

namespace XLToolbox.UnitTests.Export
{
    [TestFixture]
    class SingleExportSettingsViewModelTest
    {
        SingleExportSettingsViewModel svm;

        [SetUp]
        public void SetUp()
        {
            ExcelInstance.Start();
            svm = new SingleExportSettingsViewModel();
        }

        [TearDown]
        public void TearDown()
        {
            ExcelInstance.Shutdown();
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
        public void ChooseFileNameCommand()
        {
            bool fnMsgSent = false;
            svm.ChooseFileNameMessage.Sent +=
                (object sender, MessageArgs<StringMessageContent> args) =>
                {
                    fnMsgSent = true;
                };
            svm.ChooseFileNameCommand.Execute(null);
            Assert.IsTrue(fnMsgSent, "ChooseFileNameMessage was not sent.");
        }
    }
}
