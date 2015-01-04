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
using System.Windows.Input;

namespace Bovender.Mvvm
{
    /// <summary>
    /// Command that implements ICommand and accepts delegates
    /// that contain the command implementation.
    /// </summary>
    /// <remarks>
    /// Based on Josh Smith's article in MSDN Magazine,
    /// http://msdn.microsoft.com/en-us/magazine/dd419663.aspx
    /// </remarks>
    public class DelegatingCommand : ICommand
    {
        #region Private properties

        readonly Action<object> _execute;
        readonly Predicate<object> _canExecute;

        #endregion

        #region Constructors

        /// <summary>
        /// Creates a new command object that can always execute.
        /// </summary>
        /// <param name="execute">Code that will be executed.</param>
        public DelegatingCommand(Action<object> execute)
            : this(execute, null)
        {
        }

        public DelegatingCommand(Action<object> execute, Predicate<object> canExecute)
        {
            if (execute == null)
                throw new ArgumentNullException("execute");

            _execute = execute;
            _canExecute = canExecute;
        }

        #endregion

        #region ICommand Members

        public bool CanExecute(object parameter)
        {
            return _canExecute == null ? true : _canExecute(parameter);
        }

        public event EventHandler CanExecuteChanged
        {
            add { CommandManager.RequerySuggested += value; }
            remove { CommandManager.RequerySuggested -= value; }
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
    }
}