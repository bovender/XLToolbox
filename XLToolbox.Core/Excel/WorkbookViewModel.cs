using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Input;
using System.Collections.ObjectModel;
using System.ComponentModel;
using Microsoft.Office.Interop.Excel;

namespace XLToolbox.Core.Excel
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
        private DelegatingCommand _moveSheetUp;
        private DelegatingCommand _moveSheetDown;
        private int _numSelectedSheets = 0;

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

        private void svm_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == "IsSelected")
            {
                SheetViewModel svm = sender as SheetViewModel;
                _numSelectedSheets += (svm.IsSelected) ? 1 : -1;
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

        private bool CanMoveSheetUp()
        {
            return ((_numSelectedSheets > 0) && !Sheets[0].IsSelected);
        }

        private void DoMoveSheetDown()
        {
            throw new NotImplementedException();
        }

        private bool CanMoveSheetDown()
        {
            throw new NotImplementedException();
        }

        #endregion
    }
}
