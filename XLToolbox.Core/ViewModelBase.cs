using System.ComponentModel;

namespace XLToolbox.Core
{
    public class ViewModelBase : INotifyPropertyChanged
    {
        #region Private members

        private string _displayString;
        private bool _isSelected;

        #endregion

        #region Public properties

        public virtual string DisplayString
        {
            get
            {
                return _displayString;
            }
            set
            {
                if (value != _displayString)
                {
                    _displayString = value;
                    OnPropertyChanged("DisplayString");
                }
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

        #region INotifyPropertyChanged interface

        public event PropertyChangedEventHandler PropertyChanged;

        #endregion

        #region Protected methods

        protected virtual void OnPropertyChanged(string propertyName)
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
            }
        }
        
        #endregion

        #region Constructor

        /// <summary>
        /// Does not allow public instantiation of this class.
        /// </summary>
        protected ViewModelBase() { }

        #endregion
    }
}
