using System;
using System.ComponentModel;
using System.Windows.Input;

namespace XLToolbox.Core.Mvvm
{
    public class ViewModelBase : INotifyPropertyChanged
    {
        #region Private members

        private string _displayString;
        private DelegatingCommand _closeViewCommand;
        private bool _isSelected;

        #endregion

        #region Events

        /// <summary>
        /// Raised by the CloseView Command, signals that associated views
        /// are to be closed.
        /// </summary>
        public event EventHandler RequestCloseView;

        #endregion

        #region Commands

        public ICommand CloseViewCommand
        {
            get
            {
                if (_closeViewCommand == null)
                {
                    _closeViewCommand = new DelegatingCommand(
                        parameter => { OnCloseView(); },
                        parameter => { return CanCloseView(); }
                        );
                };
                return _closeViewCommand;
            }
        }

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

        protected virtual bool CanCloseView()
        {
            return true;
        }

        protected void OnCloseView()
        {
            if (RequestCloseView != null && CanCloseView())
            {
                RequestCloseView(this, EventArgs.Empty);
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
