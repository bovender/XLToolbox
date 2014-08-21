using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Input;
using System.Collections.ObjectModel;
using Microsoft.Office.Interop.Excel;

namespace XLToolbox.Core.Excel
{
    /// <summary>
    /// View model for an Excel workbook containing a list of sheets (worksheets, charts).
    /// </summary>
    public class WorkbookViewModel : ViewModelBase
    {
        #region Private properties

        private Workbook _workbook;
        private ObservableCollection<SheetViewModel> _sheets;

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
            foreach (dynamic sheet in Workbook.Sheets)
            {
                sheets.Add(new SheetViewModel(sheet));
            };
            this.Sheets = sheets;
        }

        #endregion
    }
}
