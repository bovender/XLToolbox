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
    public class SingleExportSettingsViewModel : SettingsViewModelBase
    {
        #region Public properties

        /// <summary>
        /// Gets or sets the desired width of the exported graphic.
        /// </summary>
        public double Width
        {
            get { return ((SingleExportSettings)Settings).Width; }
            set
            {
                ((SingleExportSettings)Settings).Width = value;
                _dimensionsChanged = true;
                OnPropertyChanged("Width");
                if (PreserveAspect) OnPropertyChanged("Height");
            }
        }

        /// <summary>
        /// Gets or sets the desired width of the exported graphic.
        /// </summary>
        public double Height
        {
            get { return ((SingleExportSettings)Settings).Height; }
            set
            {
                ((SingleExportSettings)Settings).Height = value;
                _dimensionsChanged = true;
                OnPropertyChanged("Height");
                if (PreserveAspect) OnPropertyChanged("Width");
            }
        }

        /// <summary>
        /// Preserve aspect ratio if width or height are changed.
        /// </summary>
        public bool PreserveAspect
        {
            get { return ((SingleExportSettings)Settings).PreserveAspect; }
            set
            {
                ((SingleExportSettings)Settings).PreserveAspect = value;
                OnPropertyChanged("PreserveAspect");
            }
        }

        #endregion

        #region Commands

        /// <summary>
        /// Resets the Height and Width properties to the dimensions
        /// of the current selection in Excel.
        /// </summary>
        public DelegatingCommand ResetDimensionsCommand
        {
            get
            {
                if (_resetDimensionsCommand == null)
                {
                    _resetDimensionsCommand = new DelegatingCommand(
                        param => DoResetDimensions(),
                        param => CanResetDimensions()
                    );
                }
                return _resetDimensionsCommand;
            }
        }

        /// <summary>
        /// Causes the <see cref="ChooseFileNameMessage"/> to be sent.
        /// Upon confirmation of this message, the Export process will
        /// be started.
        /// </summary>
        public DelegatingCommand ChooseFileNameCommand
        {
            get
            {
                if (_chooseFileNameCommand == null)
                {
                    _chooseFileNameCommand = new DelegatingCommand(
                        param => DoChooseFileName());
                }
                return _chooseFileNameCommand;
            }
        }

        #endregion

        #region Messages

        public Message<StringMessageContent> ChooseFileNameMessage
        {
            get
            {
                if (_chooseFileNameMessage == null)
                {
                    _chooseFileNameMessage = new Message<StringMessageContent>();
                }
                return _chooseFileNameMessage;
            }
        }

        #endregion

        #region Constructors

        public SingleExportSettingsViewModel()
            : base()
        {
            Settings = new SingleExportSettings();
        }

        /*
        /// <summary>
        /// Instantiates the view model and adds the <paramref name="presetViewModel"/>
        /// to the Presets repository.
        /// </summary>
        /// <param name="presetViewModel">Preset view model to add to the repository.</param>
        public SingleExportSettingsViewModel(PresetViewModel presetViewModel)
            : this()
        {
            PresetsRepository.Presets.Add(presetViewModel);
        }

        public SingleExportSettingsViewModel(PresetViewModel preset, double width, double height)
            : this(preset)
        {
            Width = width;
            Height = height;
            _dimensionsChanged = false;
        }

        public SingleExportSettingsViewModel(PresetViewModel preset, double width, double height, bool preserveAspect)
            : this(preset, width, height)
        {
            PreserveAspect = preserveAspect;
        }

        public SingleExportSettingsViewModel(PresetViewModel preset, SelectionViewModel selection, bool preserveAspect)
            : this(preset)
        {
            Height = selection.Bounds.Height;
            Width = selection.Bounds.Width;
            _dimensionsChanged = false;
            PreserveAspect = preserveAspect;
        }
        */

        #endregion

        #region Implementation of SettingsViewModelBase

        /// <summary>
        /// Determins the suggested target directory and sends the
        /// ChooseFileNameMessage.
        /// </summary>
        protected override void DoExport()
        {
            if (CanExport())
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

        protected override bool CanExport()
        {
            if (ExcelInstance.Running)
            {
                SelectionViewModel svm = new SelectionViewModel(ExcelInstance.Application);
                return svm.Selection != null;
            }
            else
            {
                return false;
            }
        }

        #endregion

        #region Private methods

        private void DoChooseFileName()
        {
            ChooseFileNameMessage.Send(
                new StringMessageContent(GetExportPath()),
                (content) => DoConfirmFileName(content)
            );
        }

        /// <summary>
        /// Called by Message.Respond() if the user has confirmed a file name
        /// in a view subscribed to the ChooseFileNameMessage. Triggers the
        /// actual export with the file name contained in the message content.
        /// </summary>
        /// <param name="messageContent"></param>
        private void DoConfirmFileName(StringMessageContent messageContent)
        {
            if (messageContent.Confirmed)
            {
                ((SingleExportSettings)Settings).FileName = messageContent.Value;
                DoExport();
            }
        }

        private void DoResetDimensions()
        {
            if (CanResetDimensions())
            {
                SelectionViewModel selection = new SelectionViewModel(
                    Excel.Instance.ExcelInstance.Application);
                bool oldAspectSwitch = PreserveAspect;
                PreserveAspect = false;
                Width = selection.Bounds.Width;
                Height = selection.Bounds.Height;
                PreserveAspect = oldAspectSwitch;
                _dimensionsChanged = false;
            }
        }

        private bool CanResetDimensions()
        {
            return _dimensionsChanged;
        }

        #endregion

        #region Private fields

        DelegatingCommand _chooseFileNameCommand;
        DelegatingCommand _resetDimensionsCommand;
        bool _dimensionsChanged;
        private Message<StringMessageContent> _chooseFileNameMessage;

        #endregion
    }
}
