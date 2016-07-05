/* CsvFileViewModel.cs
 * part of Daniel's XL Toolbox NG
 * 
 * Copyright 2014-2016 Daniel Kraus
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
using System.Diagnostics;

namespace XLToolbox.Csv
{
    public class CsvImportViewModel : ProcessViewModelBase
    {
        #region Factory

        public static CsvImportViewModel FromLastUsed()
        {
            return new CsvImportViewModel(CsvImporter.LastImport());
        }

        #endregion

        #region Public properties

        public string FileName
        {
            get { return Importer.FileName; }
            set
            {
                Importer.FileName = value;
                OnPropertyChanged("FileName");
            }
        }

        public string FieldSeparator
        {
            get { return Importer.FieldSeparator; }
            set
            {
                Importer.FieldSeparator = value;
                OnPropertyChanged("FieldSeparator");
            }
        }

        public string DecimalSeparator
        {
            get { return Importer.DecimalSeparator; }
            set
            {
                Importer.DecimalSeparator = value;
                OnPropertyChanged("DecimalSeparator");
            }
        }

        public string ThousandsSeparator
        {
            get { return Importer.ThousandsSeparator; }
            set
            {
                Importer.ThousandsSeparator = value;
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

        public Message<FileNameMessageContent> ChooseImportFileNameMessage
        {
            get
            {
                if (_chooseImportFileNameMessage == null)
                {
                    _chooseImportFileNameMessage = new Message<FileNameMessageContent>();
                }
                return _chooseImportFileNameMessage;
            }
        }

        #endregion

        #region Constructors

        public CsvImportViewModel()
            : this(new CsvImporter()) { }

        protected CsvImportViewModel(CsvImporter model)
            : base(model)
        { }

        #endregion

        #region ProcessViewModelBase implementation

        protected override int GetPercentCompleted()
        {
            return 50; // TODO
        }

        #endregion

        #region Private methods

        private void DoChooseFileName()
        {
            WorkbookStorage.Store store = new WorkbookStorage.Store();
            ChooseImportFileNameMessage.Send(
                new FileNameMessageContent(
                    store.Get("csv_path", Excel.ViewModels.Instance.Default.ActivePath),
                    "CSV files|*.csv;*.txt;*.dat|All files|*.*"),
                ConfirmChooseFileName);
        }

        private void ConfirmChooseFileName(FileNameMessageContent messageContent)
        {
            if (messageContent.Confirmed)
            {
                Importer.FileName = messageContent.Value;
                DoImport();
            }
        }

        private void DoImport()
        {
            using (WorkbookStorage.Store store = new WorkbookStorage.Store())
            {
                store.Put("csv_path", System.IO.Path.GetDirectoryName(FileName));
            }
            Importer.Execute();
            CloseViewCommand.Execute(null);
        }

        #endregion

        #region Private properties

        private CsvImporter Importer
        {
            [DebuggerStepThrough]
            get
            {
                return ProcessModel as CsvImporter;
            }
        }
        #endregion

        #region Private fields

        DelegatingCommand _chooseFileNameCommand;
        DelegatingCommand _importCommand;
        Message<FileNameMessageContent> _chooseImportFileNameMessage;

        #endregion
    }
}
