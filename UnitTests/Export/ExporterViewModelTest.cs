using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NUnit.Framework;
using XLToolbox.Export;
using XLToolbox.Excel.Instance;
using Bovender.Mvvm.Messaging;

namespace XLToolbox.UnitTests.Export
{
    [TestFixture]
    public class ExporterViewModelTest
    {
        [Test]
        public void ExportCommandDisabledWithoutSelection()
        {
            ExporterViewModel evm = new ExporterViewModel();
            Assert.IsFalse(evm.ExportCommand.CanExecute(null),
                "Export command should be disabled if there is no selection.");
            using (new ExcelInstance())
            {
                ExcelInstance.CreateWorkbook();
                Assert.IsTrue(evm.ExportCommand.CanExecute(null),
                    "Export command should be enabled if something is selected.");
            }
        }

        [Test]
        public void ExportCommand()
        {
            using (new ExcelInstance())
            {
                ExcelInstance.CreateWorkbook();
                ExporterViewModel evm = new ExporterViewModel();
                bool fnMsgSent = false;
                evm.ChooseFileNameMessage.Sent +=
                    (object sender, MessageArgs<StringMessageContent> args) =>
                    {
                        fnMsgSent = true;
                        // args.Respond();
                    };
                evm.ExportCommand.Execute(null);
                Assert.IsTrue(fnMsgSent, "ChooseFileNameMessage was not sent.");
            }
        }

        [Test]
        public void EditSettings()
        {
            ExporterViewModel evm = new ExporterViewModel();
            bool messageSent = false;
            evm.EditSettingsMessage.Sent += (object sender, MessageArgs<ViewModelMessageContent> args) =>
            {
                messageSent = true;
                Assert.IsTrue(args.Content.ViewModel is PresetsRepositoryViewModel);
            };
            evm.EditSettingsCommand.Execute(null);
            Assert.IsTrue(messageSent, "EditSettingsMessage was not sent.");
        }
    }
}
