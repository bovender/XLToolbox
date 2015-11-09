/* CsvExportViewModel.cs
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
using Bovender.Mvvm.ViewModels;
using Bovender.Mvvm.Messaging;
using Bovender.Mvvm;
using Microsoft.Office.Interop.Excel;
using System.Threading.Tasks;
using System.Threading;

namespace XLToolbox.Csv
{
    class CsvExportViewModel : ViewModelBase
    {
        #region Factory

        public static CsvExportViewModel FromLastUsed()
        {
            return new CsvExportViewModel(CsvFile.LastExport());
        }

        #endregion

        #region Public properties

        public string FileName
        {
            get { return _csvFile.FileName; }
            set
            {
                _csvFile.FileName = value;
                OnPropertyChanged("FileName");
            }
        }

        public string FieldSeparator
        {
            get { return _csvFile.FieldSeparator; }
            set
            {
                _csvFile.FieldSeparator = value;
                OnPropertyChanged("FieldSeparator");
            }
        }

        public string DecimalSeparator
        {
            get { return _csvFile.DecimalSeparator; }
            set
            {
                _csvFile.DecimalSeparator = value;
                OnPropertyChanged("DecimalSeparator");
            }
        }

        public string ThousandsSeparator
        {
            get { return _csvFile.ThousandsSeparator; }
            set
            {
                _csvFile.ThousandsSeparator = value;
                OnPropertyChanged("ThousandsSeparator");
            }
        }

        /// <summary>
        /// Gets or sets the range to be exported.
        /// </summary>
        public Range Range { get; set; }

        #endregion

        #region Commands

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

        public DelegatingCommand ExportCommand
        {
            get
            {
                if (_exportCommand == null)
                {
                    _exportCommand = new DelegatingCommand(
                        param => DoExport());
                }
                return _exportCommand;
            }
        }

        #endregion

        #region Messages

        public Message<FileNameMessageContent> ChooseExportFileNameMessage
        {
            get
            {
                if (_chooseExportFileNameMessage == null)
                {
                    _chooseExportFileNameMessage = new Message<FileNameMessageContent>();
                }
                return _chooseExportFileNameMessage;
            }
        }

        public Message<ProcessMessageContent> ShowExportProgress
        {
            get
            {
                if (_showExportProgress == null)
                {
                    _showExportProgress = new Message<ProcessMessageContent>();
                }
                return _showExportProgress;
            }
        }

        public Message<StringMessageContent> ExportFailedMessage
        {
            get
            {
                if (_exportFailedMessage == null)
                {
                    _exportFailedMessage = new Message<StringMessageContent>();
                }
                return _exportFailedMessage;
            }
        }

        #endregion

        #region Constructors

        public CsvExportViewModel()
            : this(new CsvFile()) { }

        protected CsvExportViewModel(CsvFile model)
            : base()
        {
            _csvFile = model;
            _csvFile.ExportProgressCompleted += CsvFile_ExportProgressCompleted;
            _csvFile.ExportFailed += CsvFile_ExportFailed;
        }

        #endregion

        #region ViewModelBase implementation

        public override object RevealModelObject()
        {
            return _csvFile;
        }

        #endregion

        #region Private methods

        private void DoChooseFileName()
        {
            WorkbookStorage.Store store = new WorkbookStorage.Store();
            ChooseExportFileNameMessage.Send(
                new FileNameMessageContent(
                    store.Get("csv_path", Excel.ViewModels.Instance.Default.ActiveWorkbook.Path),
                    "CSV files|*.csv;*.txt;*.dat|All files|*.*"),
                ConfirmChooseFileName);
        }

        private void ConfirmChooseFileName(FileNameMessageContent messageContent)
        {
            if (messageContent.Confirmed)
            {
                _csvFile.FileName = messageContent.Value;
                DoExport();
            }
        }

        private void DoExport()
        {
            WorkbookStorage.Store store = new WorkbookStorage.Store();
            store.Put("csv_path", System.IO.Path.GetDirectoryName(FileName));
            _progressTimer = new Timer(UpdateProgress, null, 1000, 300);
            if (Range != null)
            {
                _csvFile.Export(Range);
            }
            else
            {
                _csvFile.Export();
            }
            CloseViewCommand.Execute(null);
        }

        void CsvFile_ExportProgressCompleted(object sender, EventArgs e)
        {
            ExportProcessMessageContent.CompletedMessage.Send();
        }

        void CsvFile_ExportFailed(object sender, System.IO.ErrorEventArgs e)
        {
            ExportFailedMessage.Send(
                new StringMessageContent(
                    String.Format(Strings.CsvExportFailed,
                    e.GetException().Message)));
        }

        void UpdateProgress(object state)
        {
            if (!_showExportProgressWasSent)
            {
                _showExportProgressWasSent = true;
                if (_csvFile.IsProcessing)
                {
                    ShowExportProgress.Send(ExportProcessMessageContent);
                }
            }
            if (_csvFile.IsProcessing)
            {
                ExportProcessMessageContent.PercentCompleted =
                   Convert.ToInt32(100d * _csvFile.CellsProcessed / _csvFile.CellsTotal);
            }
            else
            {
                _progressTimer.Dispose();
            }
        }

        #endregion

        #region Private properties

        ProcessMessageContent ExportProcessMessageContent
        {
            get
            {
                if (_exportProcessMessageContent == null)
                {
                    _exportProcessMessageContent = new ProcessMessageContent(() => _csvFile.CancelExport());
                }
                return _exportProcessMessageContent;
            }
        }

        #endregion

        #region Private fields

        CsvFile _csvFile;
        DelegatingCommand _chooseFileNameCommand;
        DelegatingCommand _exportCommand;
        Message<FileNameMessageContent> _chooseExportFileNameMessage;
        ProcessMessageContent _exportProcessMessageContent;
        Message<ProcessMessageContent> _showExportProgress;
        Message<StringMessageContent> _exportFailedMessage;
        Timer _progressTimer;
        bool _showExportProgressWasSent;

        #endregion
    }
}
