using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.ComponentModel;
using Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;
using Bovender.Mvvm.ViewModels;

namespace XLToolbox.Excel.ViewMmodels
{
    /// <summary>
    /// A view model for Excel sheets (worksheets, charts).
    /// </summary>
    public class SheetViewModel : ViewModelBase
    {
        #region Private members

        private dynamic _sheet;

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
                if (!IsValidName(value))
                {
                    throw new InvalidSheetNameException(
                        String.Format("The string '{0}' is not a valid sheet name",
                        value)
                    );
                };
                _sheet.Name = value;
                base.DisplayString = value;
            }
        }

        public dynamic Sheet
        {
            get
            {
                return _sheet;
            }
            set
            {
                _sheet = value;
                OnPropertyChanged("Sheet");
                // Set the base class' DisplayString property to prevent
                // renaming the worksheet that is triggered by writing this
                // class' DisplayString property.
                base.DisplayString = _sheet.Name;
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

        #region Static Methods

        /// <summary>
        /// Tests whether a string represents a valid Excel sheet name.
        /// </summary>
        /// <remarks>Excel sheet names must be 1 to 31 characters long and must
        /// not contain the characters ":/\[]*?".</remarks>
        /// <param name="name">String to test.</param>
        /// <returns>True if <paramref name="name"/> can be used as a sheet name,
        /// false if not.</returns>
        public static bool IsValidName(string name)
        {
            if (!String.IsNullOrEmpty(name))
            {
                Regex r = new Regex(@"^[^:/\\*?[\]]{1,31}$");
                return r.IsMatch(name);
            }
            else
            {
                return false;
            }
        }

        #endregion
    }
}
