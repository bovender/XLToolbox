using Bovender.Mvvm.ViewModels;
/* DelegatingCommand.cs
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
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq.Expressions;
using System.Windows.Input;
using System.Windows.Threading;

namespace Bovender.Mvvm
{
    /// <summary>
    /// Command that implements ICommand and accepts delegates
    /// that contain the command implementation.
    /// </summary>
    /// <remarks>
    /// Based on Josh Smith's article in MSDN Magazine,
    /// http://msdn.microsoft.com/en-us/magazine/dd419663.aspx.
    /// ListenOn method inspired by http://stackoverflow.com/a/1857619/270712 by Tomas Kafka
    /// http://stackoverflow.com/users/38729/tom%C3%A1%C5%A1-kafka (CC BY-SA 3.0)
    /// </remarks>
    public class DelegatingCommand : ICommand
    {
        #region Constructors

        /// <summary>
        /// Creates a new command object that can always execute.
        /// </summary>
        /// <param name="execute">Code that will be executed.</param>
        public DelegatingCommand(Action<object> execute)
            : this(execute, canExecute: null)
        {
        }

        /// <summary>
        /// Creates a new command object that knows the view model that it belongs to.
        /// </summary>
        /// <param name="execute">Execute method.</param>
        /// <param name="viewModel">View model that this command belongs to.</param>
        public DelegatingCommand(Action<object> execute, ViewModelBase viewModel)
            : this(execute)
        {
            _viewModel = viewModel;
        }

        /// <summary>
        /// Creates a new command object whose executable state is determined by the
        /// <paramref name="canExecute"/> method.
        /// </summary>
        /// <param name="execute">Execute method.</param>
        /// <param name="canExecute">Function that determines whether the command can
        /// be executed or not.</param>
        public DelegatingCommand(Action<object> execute, Predicate<object> canExecute)
        {
            if (execute == null)
                throw new ArgumentNullException("execute");

            _execute = execute;
            _canExecute = canExecute;
        }

        /// <summary>
        /// Creates a new command object whose executable state is determined by the
        /// <paramref name="canExecute"/> method and that know the view model that it
        /// belongs to.
        /// </summary>
        /// <param name="execute">Execute method.</param>
        /// <param name="canExecute">Function that determines whether the command can
        /// be executed or not.</param>
        /// <param name="viewModel">View model that this command belongs to.</param>
        public DelegatingCommand(Action<object> execute, Predicate<object> canExecute,
             ViewModelBase viewModel)
            : this(execute, canExecute)
        {
            _viewModel = viewModel;
        }

        #endregion

        #region ICommand Members

        public bool CanExecute(object parameter)
        {
            return _canExecute == null ? true : _canExecute(parameter);
        }

        public event EventHandler CanExecuteChanged
        {
            add
            {
                CommandManager.RequerySuggested += value;
                CommandManagerHelper.AddWeakReferenceHandler(ref _canExecuteChangedHandlers, value, 2);
            }
            remove
            {
                CommandManager.RequerySuggested -= value;
                CommandManagerHelper.RemoveWeakReferenceHandler(_canExecuteChangedHandlers, value);
            }
        }

        public void Execute(object parameter)
        {
            try
            {
                _execute(parameter);
            }
            catch (Exception e)
            {
                // Let the central exception manager do its work.
                // If the exception was not managed, rethrow it.
                if (!ExceptionHandler.CentralHandler.Manage(this, e))
                {
                    throw;
                }
            }
        }

        #endregion // ICommand Members

        #region Public methods

        /// <summary>
        /// Makes DelegateCommnand listen on PropertyChanged events of some object,
        /// so that DelegateCommnand can update its IsEnabled property.
        /// </summary>
        public DelegatingCommand ListenOn(string propertyName)
        {
            if (_viewModel == null)
            {
                throw new InvalidOperationException(
                    "To execute 'ListenOn', construct this DelegatingCommand with the parent view model.");
            }
            _viewModel.PropertyChanged += (object sender, PropertyChangedEventArgs e) =>
            {
                if (e.PropertyName == propertyName)
                {
                    if (_viewModel.ViewDispatcher != null)
                    {
                        _viewModel.ViewDispatcher.BeginInvoke(
                            new Action(
                                () => this.OnCanExecuteChanged()
                                )
                            );
                    }
                    else
                    {
                        this.OnCanExecuteChanged();
                    }
                }
            };

            return this;
        }

        #endregion

        #region Protected methods

        protected virtual void OnCanExecuteChanged()
        {
            CommandManagerHelper.CallWeakReferenceHandlers(_canExecuteChangedHandlers);
        }

        #endregion

        #region Private properties

        readonly Action<object> _execute;
        readonly Predicate<object> _canExecute;
        readonly ViewModelBase _viewModel;
        private List<WeakReference> _canExecuteChangedHandlers;

        #endregion
    }
}