using System;
using System.Windows;
using System.Windows.Interactivity;
using Bovender.Mvvm.Messaging;

namespace Bovender.Mvvm.Actions
{
    /// <summary>
    /// Abstract base class for MVVM messaging actions. Derived classes must
    /// implement a CreateView method that returns a view for the view model
    /// that is expected to be received as a message content.
    /// </summary>
    public abstract class MessageActionBase : TriggerAction<FrameworkElement>
    {
        #region Public properties

        public string Caption { get; set; }
        public string Message { get; set; }
        public MessageContent Content { get; private set; }

        #endregion

        #region TriggerAction overrides

        protected override void Invoke(object parameter)
        {
            dynamic args = parameter;
            if (args != null)
            {
                Content = args.Content;
                Window window = CreateView();
                // Only set the window's DataContext if it has not already been
                // assigned.
                if (window.DataContext == null)
                {
                    window.DataContext = this;
                }
                EventHandler closeHandler = null;
                closeHandler = (sender, e) =>
                {
                    Content.RequestCloseView -= closeHandler;
                    window.Close();
                    args.Respond();
                };
                Content.RequestCloseView += closeHandler;
                window.ShowDialog();
            }
        }

        #endregion

        #region Abstract methods

        /// <summary>
        /// Returns a view that can bind to expected message contents.
        /// </summary>
        /// <returns>Descendant of Window.</returns>
        protected abstract Window CreateView();

        #endregion
    }
}
