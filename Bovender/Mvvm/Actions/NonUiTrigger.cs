using System.Windows;
using System.Windows.Interactivity;

namespace Bovender.Mvvm.Actions
{
    /// <summary>
    /// A trigger that may be invoked from code.
    /// </summary>
    /// <remarks>
    /// After http://stackoverflow.com/a/12977944/270712
    /// </remarks>
    class NonUiTrigger : TriggerBase<DependencyObject>
    {
        /// <summary>
        /// Invokes the trigger's actions.
        /// </summary>
        /// <param name="parameter">The parameter value.</param>
        public void Invoke(object parameter)
        {
            this.InvokeActions(parameter);
        }
    }
}
