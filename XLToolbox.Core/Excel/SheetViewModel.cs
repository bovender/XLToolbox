using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.ComponentModel;
using Microsoft.Office.Interop.Excel;

namespace XLToolbox.Core.Excel
{
    /// <summary>
    /// A view model for Excel sheets (worksheets, charts).
    /// </summary>
    public class SheetViewModel : ViewModelBase
    {
        #region Private members

        private dynamic _sheet;
        private bool _isSelected;

        #endregion

        #region Public properties

        public override string DisplayString
        {
            get
            {
                return base.DisplayString;
            }
            set
            {
                // Todo: Make sure this does not throw a COM exception if invalid name is used
                _sheet.Name = value;
                base.DisplayString = value;
            }
        }

        public object Sheet
        {
            get
            {
                return _sheet;
            }
            set
            {
                _sheet = value;
                OnPropertyChanged("Sheet");
                this.DisplayString = _sheet.Name;
            }
        }

        public bool IsSelected
        {
            get
            {
                return _isSelected;
            }
            set
            {
                _isSelected = value;
                OnPropertyChanged("IsSelected");
            }
        }

        #endregion

        #region Constructors

        public SheetViewModel() {}

        public SheetViewModel(object sheet)
            : this()
        {
            this.Sheet = sheet;
        }

        #endregion
    }
}
