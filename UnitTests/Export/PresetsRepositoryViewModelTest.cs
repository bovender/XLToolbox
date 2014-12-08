using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NUnit.Framework;
using Bovender.Mvvm.Messaging;
using XLToolbox.Export;

namespace XLToolbox.UnitTests.Export
{
    [TestFixture]
    public class PresetsRepositoryViewModelTest
    {
        [Test]
        public void AddCommand()
        {
            PresetsRepositoryForTesting sr = new PresetsRepositoryForTesting();
            PresetsRepositoryViewModel srvm = new PresetsRepositoryViewModel(sr);
            int numSettings = sr.ExportSettings.Count;
            srvm.AddSettingsCommand.Execute(null);
            Assert.AreEqual(numSettings + 1, sr.ExportSettings.Count,
                "Export settings repository should have new settings.");
            Assert.AreEqual(numSettings + 1, srvm.ExportSettings.Count,
                "Settings repository view model should have new settings view model.");
        }

        [Test]
        public void RemoveCommand()
        {
            PresetsRepositoryForTesting sr = new PresetsRepositoryForTesting();
            for (int i = 0; i < 3; i++)
            {
                sr.Add(new Preset());
            };
            int oldCount = sr.ExportSettings.Count;
            PresetsRepositoryViewModel srvm = new PresetsRepositoryViewModel(sr);
            Preset s = sr.ExportSettings[2];
            Assert.IsFalse(srvm.RemoveSettingsCommand.CanExecute(null),
                "Remove command should be disabled if no export settings objects are selected.");
            srvm.ExportSettings[2].IsSelected = true;
            Assert.IsTrue(srvm.RemoveSettingsCommand.CanExecute(null),
                "Remove command should be enabled if at least one export settings object is selected.");
            bool messageSent = false;
            srvm.ConfirmRemoveMessage.Sent += (object sender, MessageArgs<MessageContent> args) =>
            {
                messageSent = true;
                args.Content.Confirmed = true;
                args.Respond();
            };
            srvm.RemoveSettingsCommand.Execute(null);
            Assert.IsTrue(messageSent, "ConfirmRemoveMessage was not sent.");
            srvm.RemoveSettingsCommand.Execute(null);
            Assert.AreEqual(oldCount - 1, srvm.ExportSettings.Count,
                "Number of view model messages was not reduced by 1 after delete command.");
            Assert.IsFalse(((PresetsRepository)srvm.RevealModelObject()).ExportSettings.Contains(s),
                "Settings object was supposed to be removed but still exists in SettingsRepository.");
        }

        [Test]
        public void EditCommand()
        {
            Preset s = new Preset() { Name = "test settings" };
            PresetsRepositoryForTesting sr = new PresetsRepositoryForTesting();
            PresetsRepositoryViewModel srvm = new PresetsRepositoryViewModel(sr);
            sr.Add(s);
            Assert.IsFalse(srvm.EditSettingsCommand.CanExecute(null),
                "Edit settings command should be disabled if nothing is selected.");
            srvm.ExportSettings[0].IsSelected = true;
            Assert.IsTrue(srvm.EditSettingsCommand.CanExecute(null),
                "Edit settings command should be enabled if at least one object is selected.");
            bool messageSent = false;
            srvm.EditSettingsMessage.Sent += (object sender, MessageArgs<ViewModelMessageContent> args) =>
                {
                    messageSent = true;
                    Assert.IsTrue(args.Content.ViewModel.IsViewModelOf(s),
                        "EditSettingsMessage did not carry the correct ExportSettingsViewModel object.");
                };
            srvm.EditSettingsCommand.Execute(null);
            Assert.IsTrue(messageSent, "EditSettingsMessage should have been sent but wasn't.");
        }
    }
}
