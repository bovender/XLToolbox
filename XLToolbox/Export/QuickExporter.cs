using System;
using Bovender.Mvvm.Messaging;
using Bovender.Mvvm.Actions;
using XLToolbox.Export.Models;
using XLToolbox.Export.ViewModels;

namespace XLToolbox.Export
{
    /// <summary>
    /// Provides user-entry points for the 'quick' export
    /// functions that re-use previously used settings.
    /// </summary>
    public class QuickExporter
    {
        #region Public methods

        /// <summary>
        /// Exports the current selection using the last settings, if available.
        /// </summary>
        public void ExportSelection()
        {
            // Todo: Make this check if export is possible at all.
            PresetViewModel pvm = Properties.Settings.Default.ExportPresetViewModel;
            if (pvm == null)
            {
                Dispatcher.Execute(Command.ExportSelection);
            }
            else
            {
                SingleExportSettingsViewModel svm = new SingleExportSettingsViewModel();
                svm.SelectedPreset = pvm;
                svm.ChooseFileNameMessage.Sent += ChooseFileNameMessage_Sent; 
                if (svm.ChooseFileNameCommand.CanExecute(null))
                {
                    svm.ChooseFileNameCommand.Execute(null);
                    svm.ExportCommand.Execute(null);
                }
            }
        }

        /// <summary>
        /// Performs a batch export using the last used settings, if available.
        /// </summary>
        public void ExportBatch()
        {
            // Todo: Check if export is possible at all.
            BatchExportSettings settings = Properties.Settings.Default.LastBatchExportSetting;
            if (settings == null)
            {
                Dispatcher.Execute(Command.BatchExport);
            }
            else
            {
                throw new NotImplementedException();
            }
        }

        #endregion

        #region Private methods

        void ChooseFileNameMessage_Sent(object sender, MessageArgs<FileNameMessageContent> e)
        {
            ChooseFileSaveAction action = new ChooseFileSaveAction();
            action.Invoke(e);
        }

        #endregion
    }
}
