using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using Bovender.Mvvm;
using Bovender.Mvvm.Messaging;
using Bovender.Mvvm.ViewModels;
using XLToolbox.WorkbookStorage;

namespace XLToolbox.Export
{
    /// <summary>
    /// Abstract base class for the <see cref="SingleExportSettingsViewModel"/>
    /// and <see cref="BatchExportSettingsViewModel"/> classes.
    /// </summary>
    public abstract class SettingsViewModelBase : ViewModelBase
    {
        #region Public properties

        /// <summary>
        /// Preset to use for the graphic export.
        /// </summary>
        public PresetViewModel Preset
        {
            get
            {
                if (_preset == null)
                {
                    _preset = new PresetViewModel();
                }
                return _preset;
            }
            set
            {
                _preset = value;
                OnPropertyChanged("Preset");
            }
        }

        public string FileName
        {
            get { return Settings.FileName; }
            set
            {
                Settings.FileName = value;
                OnPropertyChanged("FileName");
            }
        }
        
        #endregion

        #region Protected properties 

        protected Settings Settings { get; set; }

        #endregion

        #region Commands

        public DelegatingCommand EditPresetCommand
        {
            get
            {
                if (_editPresetCommand == null)
                {
                    _editPresetCommand = new DelegatingCommand(param => DoEditPreset());
                }
                return _editPresetCommand;
            }
        }

        public DelegatingCommand ExportCommand
        {
            get
            {
                if (_exportCommand == null)
                {
                    _exportCommand = new DelegatingCommand(
                        param => DoExport(),
                        param => CanExport()
                    );
                }
                return _exportCommand;
            }
        }

        #endregion

        #region Messages

        public Message<ViewModelMessageContent> EditPresetMessage
        {
            get
            {
                if (_editPresetMessage == null)
                {
                    _editPresetMessage = new Message<ViewModelMessageContent>();
                }
                return _editPresetMessage;
            }
        }

        public Message<ProcessMessageContent> ExportProcessMessage
        {
            get
            {
                if (_exportProcessMessage == null)
                {
                    _exportProcessMessage = new Message<ProcessMessageContent>();
                }
                return _exportProcessMessage;
            }
        }

        #endregion

        #region Abstract methods

        /// <summary>
        /// Called when the ExportCommand is executed, triggers the
        /// export process.
        /// </summary>
        protected abstract void DoExport();

        protected abstract bool CanExport();

        #endregion

        #region Protected methods

        /// <summary>
        /// Returns the default path for export; this is either the previously saved
        /// path, or the path of the current workbook, or the path for 'My Documents'.
        /// </summary>
        /// <returns>Default export path.</returns>
        protected string GetExportPath()
        {
            Workbook wb = Excel.Instance.ExcelInstance.Application.ActiveWorkbook;
            Store store = new Store(wb);
            string defaultPath;
            if (wb != null && !String.IsNullOrEmpty(wb.Path))
            {
                defaultPath = System.IO.Path.GetDirectoryName(wb.Path);
            }
            else
            {
                defaultPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            }
            return store.Get("export_path", defaultPath);
        }

        #endregion

        #region Implementation of ViewModelBase

        public override object RevealModelObject()
        {
            return Settings;
        }

        #endregion

        #region Private methods

        private void DoEditPreset()
        {
            EditPresetMessage.Send(
                new ViewModelMessageContent(Preset),
                content => OnPropertyChanged("Preset")
            );
        }

        #endregion

        #region Private fields
        
        DelegatingCommand _exportCommand;
        DelegatingCommand _editPresetCommand;
        Message<ViewModelMessageContent> _editPresetMessage;
        Message<ProcessMessageContent> _exportProcessMessage;
        PresetViewModel _preset;

        #endregion
    }
}
