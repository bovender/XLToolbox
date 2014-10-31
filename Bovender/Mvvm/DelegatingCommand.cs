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
            _execute(parameter);
        }

        #endregion // ICommand Members
    }
}