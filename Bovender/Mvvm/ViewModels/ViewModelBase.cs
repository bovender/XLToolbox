/* ViewModelBase.cs
 * part of Daniel's XL Toolbox NG
 * 
 * Copyright 2014-2015 Daniel Kraus
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
using System.ComponentModel;
using System.Windows;
using System.Windows.Input;
using System.Threading;
using System.Windows.Threading;

namespace Bovender.Mvvm.ViewModels
{
    public abstract class ViewModelBase : INotifyPropertyChanged
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

        public Dispatcher ViewDispatcher { get; set; }

        #endregion

        #region Public methods

        /// <summary>
        /// Determines whether the current object is a view model
        /// of a particular model object. Returns false if either
        /// the <see cref="model"/> or the viewmodel's wrapped
        /// model object is null.
        /// </summary>
        /// <param name="model">The model to check.</param>
        /// <returns>True if <see cref="model"/> is wrapped by
        /// this; false if not (including null objects).</returns>
        public bool IsViewModelOf(object model)
        {
            object wrappedObject = RevealModelObject();
            if (model == null || wrappedObject == null)
            {
                return false;
            }
            else
            {
                return wrappedObject.Equals(model);
            }
        }

        #endregion

        #region Public (abstract) methods

        /// <summary>
        /// Returns the model object that this view model wraps or null
        /// if there is no wrapped model object.
        /// </summary>
        /// <remarks>
        /// This is a method rather than a property to make data binding
        /// more difficult (if not impossible), because binding directly
        /// to the model object is discouraged. However, certain users
        /// such as a ViewModelCollection might need access to the wrapped
        /// model object.
        /// </remarks>
        /// <returns>Model object.</returns>
        public abstract object RevealModelObject();

        #endregion

        #region Protected properties

        /// <summary>
        /// Captures the dispatcher of the thread that the
        /// object was created in.
        /// </summary>
        protected Dispatcher Dispatcher { get; private set; }

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
        protected ViewModelBase()
        {
            // Capture the current dispatcher to enable
            // asynchronous operations that a view can
            // react to dispite running in another thread.
            Dispatcher = Dispatcher.CurrentDispatcher;
        }

        #endregion

        #region Injectors

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
            if (view != null)
            {
                EventHandler h = null;
                h = (sender, args) =>
                {
                    this.RequestCloseView -= h;
                    view.DataContext = null;
                    // view.Close();
                    view.Dispatcher.Invoke(new Action(view.Close));
                };
                this.RequestCloseView += h;
                view.DataContext = this;
                ViewDispatcher = view.Dispatcher;
            }
            return view;
        }

        /// <summary>
        /// Creates a new thread that creates a new instance of the view <typeparamref name="T"/>
        /// and shows it modelessly. Use this to show views during asynchronous operations.
        /// </summary>
        /// <typeparam name="T">View (descendant of Window).</typeparam>
        public void InjectAndShowInThread<T>() where T: Window, new()
        {
            Thread t = new Thread(() =>
            {
                T view = new T();
                EventHandler h = null;
                h = (sender, args) =>
                {
                    this.RequestCloseView -= h;
                    // view.Close();
                    view.Dispatcher.Invoke(new Action(view.Close));
                };
                this.RequestCloseView += h;
                ViewDispatcher = view.Dispatcher;
                view.DataContext = this;
                view.Closed += (sender, args) => view.Dispatcher.InvokeShutdown();
                view.Show();
                System.Windows.Threading.Dispatcher.Run();
            });
            t.SetApartmentState(ApartmentState.STA);
            t.Start();
        }

        #endregion
    }
}
