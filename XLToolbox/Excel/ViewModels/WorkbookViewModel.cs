﻿using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using Microsoft.Office.Interop.Excel;
using Bovender.Mvvm;
using Bovender.Mvvm.ViewModels;
using Bovender.Mvvm.Messaging;

namespace XLToolbox.Excel.ViewModels
{
    /// <summary>
    /// View model for an Excel workbook containing a list of sheets (worksheets, charts)
    /// that can be managed (moved around, added, deleted, renamed).
    /// </summary>
    public class WorkbookViewModel : ViewModelBase
    {
        #region Private properties

        private Workbook _workbook;
        private ObservableCollection<SheetViewModel> _sheets;
        private SheetViewModel _lastSelectedSheet;
        private DelegatingCommand _moveSheetUp;
        private DelegatingCommand _moveSheetsToTop;
        private DelegatingCommand _moveSheetDown;
        private DelegatingCommand _moveSheetsToBottom;
        private DelegatingCommand _deleteSheets;
        private DelegatingCommand _renameSheet;
        private Message<MessageContent> _confirmDeleteMessage;
        private Message<StringMessageContent> _renameSheetMessage;

        #endregion

        #region Protected properties

        protected Workbook Workbook
        {
            get
            {
                return _workbook;
            }
            set
            {
                _workbook = value;
                OnPropertyChanged("Workbook");
                BuildSheetList();
                this.DisplayString = _workbook.Name;
            }
        }
        
        #endregion

        #region Public properties

        public int NumSelectedSheets { get; private set; }

        public ObservableCollection<SheetViewModel> Sheets
        {
            get
            {
                return _sheets;
            }
            protected set
            {
                _sheets = value;
                OnPropertyChanged("Sheets");
            }
        }

        public Message<MessageContent> ConfirmDeleteMessage
        {
            get
            {
                if (_confirmDeleteMessage == null)
                {
                    _confirmDeleteMessage = new Message<MessageContent>();
                };
                return _confirmDeleteMessage;
            }
        }

        public Message<StringMessageContent> RenameSheetMessage
        {
            get
            {
                if (_renameSheetMessage == null)
                {
                    _renameSheetMessage = new Message<StringMessageContent>();
                };
                return _renameSheetMessage;
            }
        }

        #endregion

        #region Commands

        public DelegatingCommand MoveSheetUp
        {
            get
            {
                if (_moveSheetUp == null)
                {
                    _moveSheetUp = new DelegatingCommand(
                        parameter => { DoMoveSheetUp(); },
                        parameter => { return CanMoveSheetUp(); }
                        );
                }
                return _moveSheetUp;
            }
        }

        public DelegatingCommand MoveSheetsToTop
        {
            get
            {
                if (_moveSheetsToTop == null)
                {
                    _moveSheetsToTop = new DelegatingCommand(
                        parameter => { DoMoveSheetsToTop(); },
                        parameter => { return CanMoveSheetsToTop(); }
                        );
                }
                return _moveSheetsToTop;
            }
        }

        public DelegatingCommand MoveSheetDown
        {
            get
            {
                if (_moveSheetDown == null)
                {
                    _moveSheetDown = new DelegatingCommand(
                        parameter => { DoMoveSheetDown(); },
                        parameter => { return CanMoveSheetDown(); }
                        );
                }
                return _moveSheetDown;
            }
        }

        public DelegatingCommand MoveSheetsToBottom
        {
            get
            {
                if (_moveSheetsToBottom == null)
                {
                    _moveSheetsToBottom = new DelegatingCommand(
                        parameter => { DoMoveSheetsToBottom(); },
                        parameter => { return CanMoveSheetsToBottom(); }
                        );
                }
                return _moveSheetsToBottom;
            }
        }

        public DelegatingCommand DeleteSheets
        {
            get
            {
                if (_deleteSheets == null)
                {
                    _deleteSheets = new DelegatingCommand(
                        parameter => { DoDeleteSheets(); },
                        parameter => { return CanDeleteSheets(); }
                        );
                }
                return _deleteSheets;
            }
        }

        public DelegatingCommand RenameSheet
        {
            get
            {
                if (_renameSheet == null)
                {
                    _renameSheet = new DelegatingCommand(
                        parameter => { DoRenameSheet(); },
                        parameter => { return CanRenameSheet(); }
                        );
                }
                return _renameSheet;
            }
        }

        #endregion

        #region Constructors

        public WorkbookViewModel() {}

        public WorkbookViewModel(Workbook workbook)
            : this()
        {
            this.Workbook = workbook;
        }

        #endregion

        #region Protected methods

        protected void BuildSheetList()
        {
            ObservableCollection<SheetViewModel> sheets = new ObservableCollection<SheetViewModel>();
            SheetViewModel svm;
            foreach (dynamic sheet in Workbook.Sheets)
            {
                svm = new SheetViewModel(sheet);
                svm.PropertyChanged += svm_PropertyChanged;
                sheets.Add(svm);
            };
            this.Sheets = sheets;
        }

        #endregion

        #region Event handlers

        private void svm_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == "IsSelected")
            {
                SheetViewModel svm = sender as SheetViewModel;
                if (svm.IsSelected)
                {
                    NumSelectedSheets++;
                    _lastSelectedSheet = svm;
                    svm.Sheet.Activate();
                }
                else
                {
                    NumSelectedSheets--;
                }
            }
        }

        #endregion

        #region Private methods

        private void DoMoveSheetUp()
        {
            // When iterating over the worksheet view models in the Sheets collection
            // as well as over the sheets collection of the workbook, keep in mind
            // that Excel workbook collections are 1-based.
            for (int i = 1; i < Sheets.Count; i++)
            {
                if (Sheets[i].IsSelected)
                {
                    Workbook.Sheets[i+1].Move(before: Workbook.Sheets[i]);
                    Sheets.Move(i, i - 1);
                }
            }
        }

        private void DoMoveSheetsToTop()
        {
            int currentTop = 0;
            for (int i = 1; i < Sheets.Count; i++)
            {
                if (Sheets[i].IsSelected)
                {
                    Workbook.Sheets[i + 1].Move(before: Workbook.Sheets[currentTop+1]);
                    Sheets.Move(i, currentTop);
                    currentTop++;
                }
            }
        }

        private bool CanMoveSheetUp()
        {
            return ((NumSelectedSheets > 0) && !Sheets[0].IsSelected);
        }

        private bool CanMoveSheetsToTop()
        {
            return CanMoveSheetUp();
        }

        private void DoMoveSheetDown()
        {
            // When iterating over the worksheet view models in the Sheets collection
            // as well as over the sheets collection of the workbook, keep in mind
            // that Excel workbook collections are 1-based.
            for (int i = Sheets.Count - 2; i > 0; i--)
            {
                if (Sheets[i].IsSelected)
                {
                    Workbook.Sheets[i + 1].Move(after: Workbook.Sheets[i + 2]);
                    Sheets.Move(i, i + 1);
                }
            }
        }

        private void DoMoveSheetsToBottom()
        {
            int currentBottom = Sheets.Count - 1;
            for (int i = currentBottom-1; i > 0; i--)
            {
                if (Sheets[i].IsSelected)
                {
                    Workbook.Sheets[i + 1].Move(after: Workbook.Sheets[currentBottom+1]);
                    Sheets.Move(i, currentBottom);
                    currentBottom--;
                }
            }
        }

        private bool CanMoveSheetDown()
        {
            return ((NumSelectedSheets > 0) && !Sheets[Sheets.Count - 1].IsSelected);
        }

        private bool CanMoveSheetsToBottom()
        {
            return CanMoveSheetDown();
        }

        private void DoDeleteSheets()
        {
            ConfirmDeleteMessage.Send(
                new MessageContent(),
                (confirmation) =>
                {
                    ConfirmDeleteSheets(confirmation);
                });
        }

        private void ConfirmDeleteSheets(MessageContent confirmation)
        {
            if (confirmation.Confirmed)
            {
                for (int i = 0; i < Sheets.Count; i++)
                {
                    if (Sheets[i].IsSelected)
                    {
                        Sheets.RemoveAt(i);
                        Workbook.Sheets[i+1].Delete();
                    }
                }
            }
        }

        private bool CanDeleteSheets()
        {
            return (NumSelectedSheets > 0);
        }

        private void DoRenameSheet()
        {
            StringMessageContent content = new StringMessageContent();
            content.Value = _lastSelectedSheet.DisplayString;
            content.Validator = (value) =>
            {
                if (SheetViewModel.IsValidName(value))
                {
                    return String.Empty;
                }
                else
                {
                    // TODO: Find a way that does not require human language here
                    return "1-21, not () /\\ [] *?";
                }
            };
            RenameSheetMessage.Send(
                content,
                (stringMessage) =>
                    {
                        ConfirmRenameSheet(stringMessage);
                    }
            );
        }

        private void ConfirmRenameSheet(StringMessageContent stringMessage)
        {
            if (CanRenameSheet() && stringMessage.Confirmed)
            {
                _lastSelectedSheet.DisplayString = stringMessage.Value;
            }
        }

        private bool CanRenameSheet()
        {
            return (NumSelectedSheets > 0);
        }

        #endregion
    }
}