using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using Bovender.Mvvm;
using Bovender.Mvvm.Messaging;
using Bovender.Mvvm.ViewModels;

namespace XLToolbox.Export
{
    /// <summary>
    /// View model for the <see cref="Exporter"/> class.
    /// </summary>
    public class ExporterViewModel : ViewModelBase
    {
        #region Public properties

        public SettingsRepositoryViewModel SettingsRepository
        {
            get { return _settings; }
            protected set
            {
                _settings = value;
                OnPropertyChanged("Settings");
            }
        }

        #endregion

        #region Constructor

        /// <summary>
        /// Creates a new view model and a new Exporter object.
        /// An ExportSettingsRepository will also be created and pre-set
        /// with a default ExportSettings object unless there are
        /// previously stored export settings.
        /// </summary>
        /// <param name="exporter"></param>
        public ExporterViewModel()
        {
            _exporter = new Exporter();

            // Define standard export settings if none exist
            SettingsRepository sr = new SettingsRepository();
            if (sr.ExportSettings.Count == 0)
            {
                sr.Add(new Settings(FileType.Png, 300, ColorSpace.Rgb));
            }
            _settings = new SettingsRepositoryViewModel(sr);
        }

        #endregion

        #region Commands

        /// <summary>
        /// Sends a ChooseFileNameMessage, whose StringMessageContent contains
        /// the last path exported to from the current workbook, or the 'My Documents'
        /// directory. Views can confirm the message and return the destination
        /// file name in the StringMessageContent's value property.
        /// </summary>
        public DelegatingCommand ExportCommand
        {
            get
            {
                if (_exportCommand == null)
                {
                    _exportCommand = new DelegatingCommand(
                        (param) => DoExport(),
                        (param) => CanExport()
                        );
                }
                return _exportCommand;
            }
        }

        public DelegatingCommand EditSettingsCommand
        {
            get
            {
                if (_editSettingsCommand == null)
                {
                    _editSettingsCommand = new DelegatingCommand(
                        (param) => DoEditSettings()
                        );
                }
                return _editSettingsCommand;
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

        public Message<ViewModelMessageContent> EditSettingsMessage
        {
            get
            {
                if (_editSettingsMessage == null)
                {
                    _editSettingsMessage = new Message<ViewModelMessageContent>();
                }
                return _editSettingsMessage;
            }
        }

        #endregion

        #region Private methods

        private void DoExport()
        {
            Workbook wb = Excel.Instance.ExcelInstance.Application.ActiveWorkbook;
            string defaultPath;
            if (wb != null && !String.IsNullOrEmpty(wb.Path))
            {
                defaultPath = System.IO.Path.GetDirectoryName(wb.Path);
            } 
            else
            {
                defaultPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            }
            string path = WorkbookStore.Get("exportpath", defaultPath);
            ChooseFileNameMessage.Send(
                new StringMessageContent(path),
                (content) => DoConfirmFileName(content)
            );
        }

        private bool CanExport()
        {
            return CurrentSelection != null;
        }

        private void DoConfirmFileName(StringMessageContent messageContent)
        {
            if (messageContent.Confirmed && CanExport())
            {
                Settings exportSettings = SettingsRepository.ExportSettings.LastSelected
                    .RevealModelObject() as Settings;
                _exporter.ExportSelection(exportSettings, messageContent.Value);
            }
        }

        private void DoEditSettings()
        {
            EditSettingsMessage.Send(
                new ViewModelMessageContent(SettingsRepository),
                (content) => OnPropertyChanged("SettingsRepository")
            );
        }

        #endregion

        #region Protected properties

        protected WorkbookStorage.Store WorkbookStore
        {
            get
            {
                if (_workbookStore == null)
                {
                    _workbookStore = new WorkbookStorage.Store();
                }
                return _workbookStore;
            }
        }

        protected object CurrentSelection
        {
            get
            {
                try
                {
                    return Excel.Instance.ExcelInstance.Application.Selection;
                }
                catch
                {
                    return null;
                }
            }
        }

        #endregion

        #region Private fields

        private Exporter _exporter;
        private DelegatingCommand _exportCommand;
        private DelegatingCommand _editSettingsCommand;
        private Message<StringMessageContent> _chooseFileNameMessage;
        private Message<ViewModelMessageContent> _editSettingsMessage;
        private SettingsRepositoryViewModel _settings;
        private WorkbookStorage.Store _workbookStore; 

        #endregion

        #region ViewModelBase implementation

        public override object RevealModelObject()
        {
            return _exporter;
        }

        #endregion
    }
}
