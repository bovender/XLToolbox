/* CsvExportViewModel.cs
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
using Microsoft.Office.Interop.Excel;
using System.Threading.Tasks;
using System.Threading;
using System.Diagnostics;

namespace XLToolbox.Csv
{
    public class CsvExportViewModel : ProcessViewModelBase
    {
        #region Factory

        public static CsvExportViewModel FromLastUsed()
        {
            return new CsvExportViewModel(CsvExporter.LastExport());
        }

        #endregion

        #region Public properties

        public string FileName
        {
            get { return CsvExporter.FileName; }
            set
            {
                CsvExporter.FileName = value;
                OnPropertyChanged("FileName");
            }
        }

        public string FieldSeparator
        {
            get { return CsvExporter.FieldSeparator; }
            set
            {
                CsvExporter.FieldSeparator = value;
                OnPropertyChanged("FieldSeparator");
            }
        }

        public string DecimalSeparator
        {
            get { return CsvExporter.DecimalSeparator; }
            set
            {
                CsvExporter.DecimalSeparator = value;
                OnPropertyChanged("DecimalSeparator");
            }
        }

        public string ThousandsSeparator
        {
            get { return CsvExporter.ThousandsSeparator; }
            set
            {
                CsvExporter.ThousandsSeparator = value;
                OnPropertyChanged("ThousandsSeparator");
            }
        }

        public bool Tabularize
        {
            get
            {
                return CsvExporter.Tabularize;
            }
            set
            {
                CsvExporter.Tabularize = value;
                OnPropertyChanged("EqualWidths");
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

        #endregion

        #region Constructors

        public CsvExportViewModel()
            : this(new CsvExporter()) { }

        public CsvExportViewModel(CsvExporter model)
            : base(model)
        { }

        #endregion

        #region Implementation of ProcessViewModel

        protected override void UpdateProcessMessageContent(ProcessMessageContent processMessageContent)
        {
            processMessageContent.PercentCompleted = Convert.ToInt32(100d * CsvExporter.CellsProcessed / CsvExporter.CellsTotal);
        }

        #endregion

        #region Private properties

        private CsvExporter CsvExporter
        {
            [DebuggerStepThrough]
            get
            {
                return ProcessModel as CsvExporter;
            }
        }
        
        #endregion

        #region Private methods

        private void DoChooseFileName()
        {
            WorkbookStorage.Store store = new WorkbookStorage.Store();
            string defaultPath = Excel.ViewModels.Instance.Default.ActivePath;
            ChooseExportFileNameMessage.Send(
                new FileNameMessageContent(store.Get("csv_path", defaultPath),
                    "CSV files|*.csv;*.txt;*.dat|All files|*.*"),
                ConfirmChooseFileName);
        }

        private void ConfirmChooseFileName(FileNameMessageContent messageContent)
        {
            if (messageContent.Confirmed)
            {
                CsvExporter.FileName = messageContent.Value;
                DoExport();
            }
        }

        private void DoExport()
        {
            using (WorkbookStorage.Store store = new WorkbookStorage.Store())
            {
                store.Put("csv_path", System.IO.Path.GetDirectoryName(FileName));
            };
            ((CsvExporter)ProcessModel).Range = Range;
            StartProcess();
        }

        #endregion

        #region Private fields

        DelegatingCommand _chooseFileNameCommand;
        DelegatingCommand _exportCommand;
        Message<FileNameMessageContent> _chooseExportFileNameMessage;

        #endregion
    }
}
