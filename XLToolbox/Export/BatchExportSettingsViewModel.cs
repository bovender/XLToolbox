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

        public string Path
        {
            get { return ((BatchExportSettings)Settings).Path; }
            set
            {
                ((BatchExportSettings)Settings).Path = value;
                OnPropertyChanged("Path");
            }
        }
       
        #endregion

        #region Commands

        /// <summary>
        /// Causes the <see cref="ChooseFolderMessage"/> to be sent.
        /// Upon confirmation of the message by a view, the Export
        /// process will be started.
        /// </summary>
        public DelegatingCommand ChooseFolderCommand
        {
            get
            {
                if (_chooseFolderCommand == null)
                {
                    _chooseFolderCommand = new DelegatingCommand(
                        param => DoChooseFolder());
                }
                return _chooseFolderCommand;
            }
        }

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
        /// Determines the suggested target directory and sends the
        /// ChooseFileNameMessage.
        /// </summary>
        private void DoChooseFolder()
        {
            ChooseFolderMessage.Send(
                new StringMessageContent(GetExportPath()),
                (content) => ConfirmFolder(content)
            );
        }

        protected override void DoExport()
        {
            if (CanExport())
            {
                // TODO: Make export asynchronous
                ProcessMessageContent pcm = new ProcessMessageContent();
                pcm.IsIndeterminate = true;
                ExportProcessMessage.Send(pcm);
                Exporter exporter = new Exporter();
                exporter.ExportBatch(Settings as BatchExportSettings);
                pcm.CompletedMessage.Send(pcm);
            }
        }

        protected override bool CanExport()
        {
            SelectionViewModel svm = new SelectionViewModel(ExcelInstance.Application);
            return svm.Selection != null;
        }

        #endregion

        #region Private methods

        private void ConfirmFolder(StringMessageContent messageContent)
        {
            if (messageContent.Confirmed)
            {
                ((BatchExportSettings)Settings).Path = messageContent.Value;
                DoExport();
            }
        }

        #endregion

        #region Private fields

        private DelegatingCommand _chooseFolderCommand;
        private Message<StringMessageContent> _chooseFolderMessage;

        #endregion
    }
}
