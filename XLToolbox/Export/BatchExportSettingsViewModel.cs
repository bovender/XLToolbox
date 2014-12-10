using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Xml;
using System.Configuration;
using Microsoft.Office.Interop.Excel;
using Bovender.Mvvm;
using Bovender.Mvvm.Messaging;
using Bovender.Mvvm.ViewModels;
using XLToolbox.Excel.ViewModels;
using XLToolbox.Excel.Instance;
using XLToolbox.WorkbookStorage;

namespace XLToolbox.Export
{
    /// <summary>
    /// View model for the <see cref="Settings"/> class.
    /// </summary>
    [Serializable]
    [SettingsSerializeAs(SettingsSerializeAs.Xml)]
    public class BatchExportSettingsViewModel : SettingsViewModelBase
    {
        #region Public properties

        public BatchExportLayout Layout
        {
            get
            {
                return ((BatchExportSettings)Settings).Layout;
            }
            set
            {
                ((BatchExportSettings)Settings).Layout = value;
                OnPropertyChanged("Layout");
            }
        }

        public BatchExportScope Scope
        {
            get
            {
                return ((BatchExportSettings)Settings).Scope;
            }
            set
            {
                ((BatchExportSettings)Settings).Scope = value;
                OnPropertyChanged("Scope");
            }
        }

        public BatchExportObjects Objects
        {
            get
            {
                return ((BatchExportSettings)Settings).Objects;
            }
            set
            {
                ((BatchExportSettings)Settings).Objects = value;
                OnPropertyChanged("Objects");
            }
        }

        #endregion

        #region Commands

        #endregion

        #region Messages

        public Message<StringMessageContent> ChooseFolderMessage
        {
            get
            {
                if (_chooseFolderMessage == null)
                {
                    _chooseFolderMessage = new Message<StringMessageContent>();
                }
                return _chooseFolderMessage;
            }
        }

        #endregion

        #region Constructors

        public BatchExportSettingsViewModel()
            : base()
        {
            Settings = new BatchExportSettings();
        }

        public BatchExportSettingsViewModel(PresetViewModel preset)
            : this()
        {
            Preset = preset;
        }

        #endregion

        #region Implementation of SettingsViewModelBase

        /// <summary>
        /// Determins the suggested target directory and sends the
        /// ChooseFileNameMessage.
        /// </summary>
        protected override void DoExport()
        {
            ChooseFolderMessage.Send(
                new StringMessageContent(GetExportPath()),
                (content) => DoConfirmFolder(content)
            );
        }

        protected override bool CanExport()
        {
            SelectionViewModel svm = new SelectionViewModel(ExcelInstance.Application);
            return svm.Selection != null;
        }

        #endregion

        #region Private methods

        /// <summary>
        /// Called by Message.Respond() if the user has confirmed a file name
        /// in a view subscribed to the ChooseFileNameMessage. Performs the
        /// actual export with the file name contained in the message content.
        /// </summary>
        /// <param name="messageContent"></param>
        private void DoConfirmFolder(StringMessageContent messageContent)
        {
            if (messageContent.Confirmed && CanExport())
            {
                // TODO: Make export asynchronous
                ProcessMessageContent pcm = new ProcessMessageContent();
                pcm.IsIndeterminate = true;
                ExportProcessMessage.Send(pcm);
                Exporter exporter = new Exporter();
                exporter.ExportSelection(Settings as SingleExportSettings);
                pcm.CompletedMessage.Send(pcm);
            }
        }

        #endregion

        #region Private fields

        private Message<StringMessageContent> _chooseFolderMessage;

        #endregion
    }
}
