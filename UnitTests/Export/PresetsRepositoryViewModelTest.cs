/* PresetsRepositoryViewModelTest.cs
 * part of Daniel's XL Toolbox NG
 * 
 * Copyright 2014-2015 Daniel Kraus
 * 
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 * 
 *     http://www.apache.org/licenses/LICENSE-2.0
 * 
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NUnit.Framework;
using Bovender.Mvvm.Messaging;
using XLToolbox.Export.Models;
using XLToolbox.Export.ViewModels;

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
            int numSettings = sr.Presets.Count;
            srvm.AddCommand.Execute(null);
            Assert.AreEqual(numSettings + 1, sr.Presets.Count,
                "Export settings repository should have new settings.");
            Assert.AreEqual(numSettings + 1, srvm.Presets.Count,
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
            int oldCount = sr.Presets.Count;
            PresetsRepositoryViewModel srvm = new PresetsRepositoryViewModel(sr);
            Preset s = sr.Presets[2];
            Assert.IsFalse(srvm.RemoveCommand.CanExecute(null),
                "Remove command should be disabled if no export settings objects are selected.");
            srvm.Presets[2].IsSelected = true;
            Assert.IsTrue(srvm.RemoveCommand.CanExecute(null),
                "Remove command should be enabled if at least one export settings object is selected.");
            bool messageSent = false;
            srvm.ConfirmRemoveMessage.Sent += (object sender, MessageArgs<MessageContent> args) =>
            {
                messageSent = true;
                args.Content.Confirmed = true;
                args.Respond();
            };
            srvm.RemoveCommand.Execute(null);
            Assert.IsTrue(messageSent, "ConfirmRemoveMessage was not sent.");
            srvm.RemoveCommand.Execute(null);
            Assert.AreEqual(oldCount - 1, srvm.Presets.Count,
                "Number of view model messages was not reduced by 1 after delete command.");
            Assert.IsFalse(((PresetsRepository)srvm.RevealModelObject()).Presets.Contains(s),
                "Settings object was supposed to be removed but still exists in SettingsRepository.");
        }

        [Test]
        public void EditCommand()
        {
            Preset s = new Preset() { Name = "test settings" };
            PresetsRepositoryForTesting sr = new PresetsRepositoryForTesting();
            PresetsRepositoryViewModel srvm = new PresetsRepositoryViewModel(sr);
            sr.Add(s);
            Assert.IsFalse(srvm.EditCommand.CanExecute(null),
                "Edit settings command should be disabled if nothing is selected.");
            srvm.Presets[srvm.Presets.Count-1].IsSelected = true;
            Assert.IsTrue(srvm.EditCommand.CanExecute(null),
                "Edit settings command should be enabled if at least one object is selected.");
            bool messageSent = false;
            srvm.EditSettingsMessage.Sent += (object sender, MessageArgs<ViewModelMessageContent> args) =>
                {
                    messageSent = true;
                    Assert.IsTrue(args.Content.ViewModel.IsViewModelOf(s),
                        "EditSettingsMessage did not carry the correct ExportSettingsViewModel object.");
                };
            srvm.EditCommand.Execute(null);
            Assert.IsTrue(messageSent, "EditSettingsMessage should have been sent but wasn't.");
        }
    }
}
