using System;
using System.Threading.Tasks;
using Bovender.Versioning;
using Bovender.Mvvm.Messaging;
using NUnit.Framework;

namespace Bovender.UnitTests
{
    [TestFixture]
    public class UpdaterViewModelTest
    {
        [Test]
        public void NoUpdateAvailable()
        {
            UpdaterForTesting updater = new UpdaterForTesting();
            UpdaterViewModel vm = new UpdaterViewModel(updater);
            bool checkFinished = false;
            bool messageSent = false;
            // Add or own event handler to updater's CheckForUpdateFinished
            // event so we know when the low-level operation has been
            // completed.
            updater.CheckForUpdateFinished += (sender, args) =>
            {
                checkFinished = true;
            };
            updater.TestVersion = "99.99.99";
            vm.NoUpdateAvailableMessage.Sent += (sender, messageArgs) =>
            {
                messageSent = true;
            };
            Task checkFinishedTask = new Task(() =>
            {
                while (checkFinished == false) ;
            });
            vm.CheckForUpdateCommand.Execute(null);
            checkFinishedTask.Start();

            // Wait for the update check to complete asynchronously
            checkFinishedTask.Wait(10000);

            Task checkMessageSentTask = new Task(() =>
            {
                while (!messageSent) ;
            });

            // Give the MVVM messaging a chance to work before
            // we assert that the message has indeed been sent.
            checkMessageSentTask.Wait(1000);
            Assert.True(messageSent, "NoUpdateAvailableMessage should have been sent but wasn't.");
        }

        [Test]
        public void UpdateAvailable()
        {
            UpdaterForTesting updater = new UpdaterForTesting();
            UpdaterViewModel vm = new UpdaterViewModel(updater);
            bool checkFinished = false;
            bool messageSent = false;
            // Add or own event handler to updater's CheckForUpdateFinished
            // event so we know when the low-level operation has been
            // completed.
            updater.CheckForUpdateFinished += (sender, args) =>
            {
                checkFinished = true;
            };
            updater.TestVersion = "0.0.0";
            vm.UpdateAvailableMessage.Sent += (sender, messageArgs) =>
            {
                messageSent = true;
            };
            Task checkFinishedTask = new Task(() =>
            {
                while (checkFinished == false) ;
            });
            vm.CheckForUpdateCommand.Execute(null);
            checkFinishedTask.Start();

            // Wait for the update check to complete asynchronously
            checkFinishedTask.Wait(10000);

            Task checkMessageSentTask = new Task(() =>
            {
                while (!messageSent) ;
            });

            // Give the MVVM messaging a chance to work before
            // we assert that the message has indeed been sent.
            checkMessageSentTask.Wait(1000);
            Assert.True(messageSent, "NoUpdateAvailableMessage should have been sent but wasn't.");

        }

        [Test]
        public void DownloadUpdate()
        {
            UpdaterForTesting updater = new UpdaterForTesting();
            updater.TestVersion = "0.0.0";
            UpdaterViewModel vm = new UpdaterViewModel(updater);
            vm.UpdateAvailableMessage.Sent += (object sender, MessageArgs<ViewModelMessageContent> args) =>
            {
                UpdaterViewModel relayedViewModel = args.Content.ViewModel as UpdaterViewModel;
                relayedViewModel.DownloadUpdateCommand.Execute(null);
            };
            bool cancelTask = false;
            bool updateInstallable = false;
            vm.UpdateInstallableMessage.Sent += (object sender, MessageArgs<ViewModelMessageContent> args) => {
                UpdaterViewModel relayedViewModel = args.Content.ViewModel as UpdaterViewModel;
                updateInstallable = true;
            };
            vm.CheckForUpdateCommand.Execute(null);

            Task checkInstallableTask = new Task(() =>
            {
                while (!updateInstallable && !cancelTask) ;
            });

            checkInstallableTask.Start();
            checkInstallableTask.Wait(10000);
            // Cancel the task in case the timeout was reached but the event was not raised
            cancelTask = !updateInstallable;
            Assert.True(updateInstallable, "Update should have been downloaded and be installable, but isn't.");
        }
    }
}
