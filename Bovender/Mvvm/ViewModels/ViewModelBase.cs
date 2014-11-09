using System;
using System.ComponentModel;
using System.Windows;
using System.Windows.Input;

namespace Bovender.Mvvm.ViewModels
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
                        parameter => { DoCloseView(); },
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

        protected virtual void DoCloseView()
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

        #region Injector

        /// <summary>
        /// Injects the ViewModel into a newly created View and wires the RequestCloseView
        /// event.
        /// </summary>
        /// <typeparam name="T">View, must be derived from <see cref="System.Windows.Window"/>
        /// </typeparam>
        /// <returns>View with DataContext set to the current ViewModel instance that
        /// responds to the RequestCloseView event by closing itself.</returns>
        public Window InjectInto<T>() where T : Window, new()
        {
            T view = new T();
            return InjectInto(view);
        }

        /// <summary>
        /// Injects the view model into an existing view by setting
        /// the view's DataContext.
        /// </summary>
        /// <param name="view">View that shall be dependency injected.</param>
        /// <returns>View with current view model injected.</returns>
        public Window InjectInto(Window view)
        {
            EventHandler h = null;
            h = (sender, args) =>
            {
                this.RequestCloseView -= h;
                // view.Close();
                view.Dispatcher.Invoke(new Action(view.Close));
            };
            this.RequestCloseView += h;
            view.DataContext = this;
            return view;
        }

        #endregion
    }
}
