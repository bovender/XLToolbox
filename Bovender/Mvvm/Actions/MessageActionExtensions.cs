using System.Windows.Interactivity;

namespace Bovender.Mvvm.Actions
{
    public static class MessageActionExtensions
    {
        /// <summary>
        /// Invokes a <see cref="TriggerAction"/> with the specified parameter.
        /// </summary>
        /// <param name="action">The <see cref="TriggerAction"/>.</param>
        /// <param name="parameter">The parameter value.</param>
        /// <remarks>
        /// After http://stackoverflow.com/a/12977944/270712
        /// </remarks>
        public static void Invoke(this MessageActionBase action, object parameter)
        {
            NonUiTrigger trigger = new NonUiTrigger();
            trigger.Actions.Add(action);

            try
            {
                trigger.Invoke(parameter);
            }
            finally
            {
                trigger.Actions.Remove(action);
            }
        }

        /// <summary>
        /// Invokes a <see cref="TriggerAction"/>.
        /// </summary>
        /// <param name="action">The <see cref="TriggerAction"/>.</param>
        public static void Invoke(this MessageActionBase action)
        {
            action.Invoke(null);
        }
    }
}