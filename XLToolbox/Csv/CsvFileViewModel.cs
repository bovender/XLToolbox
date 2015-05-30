/* CsvFileViewModel.cs
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

namespace XLToolbox.Csv
{
    class CsvFileViewModel : ViewModelBase
    {
        #region Factory

        public static CsvFileViewModel FromLastUsed()
        {
            return new CsvFileViewModel(CsvFile.FromLastUsed());
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

        public DelegatingCommand ImportCommand
        {
            get
            {
                if (_importCommand == null)
                {
                    _importCommand = new DelegatingCommand(
                        param => DoImport());
                }
                return _importCommand;
            }
        }

        #endregion

        #region Messages

        public Message<FileNameMessageContent> ChooseFileNameMessage
        {
            get
            {
                if (_chooseFileNameMessage == null)
                {
                    _chooseFileNameMessage = new Message<FileNameMessageContent>();
                }
                return _chooseFileNameMessage;
            }
        }

        #endregion

        #region Constructors

        public CsvFileViewModel()
            : this(new CsvFile()) { }

        protected CsvFileViewModel(CsvFile model)
            : base()
        {
            _csvFile = model;
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
            ChooseFileNameMessage.Send(
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
                DoImport();
            }
        }

        private void DoImport()
        {
            WorkbookStorage.Store store = new WorkbookStorage.Store();
            store.Put("csv_path", System.IO.Path.GetDirectoryName(FileName));
            _csvFile.Import();
            CloseViewCommand.Execute(null);
        }

        #endregion

        #region Private fields

        CsvFile _csvFile;
        DelegatingCommand _chooseFileNameCommand;
        DelegatingCommand _importCommand;
        Message<FileNameMessageContent> _chooseFileNameMessage;

        #endregion
    }
}
