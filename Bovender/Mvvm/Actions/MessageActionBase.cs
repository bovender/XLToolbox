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
        public MessageContent Content { get; protected set; }

        #endregion

        #region TriggerAction overrides

        /// <summary>
        /// Creates a view that has its dependent view model injected
        /// into it.
        /// </summary>
        /// <remarks>
        /// This methods delegates the creation of the view instance to
        /// the virtual <see cref="CreateView"/> method. If this method
        /// injects a view model into the view's DataContext, the view
        /// will simply be shown (as a dialog). If the CreateView method
        /// does not inject the dependeny, the Invoke method will inject
        /// 'this' (i.e., the MessageActionBase object or a descendant)
        /// into the view and sets the RequestClose event handler.
        /// </remarks>
        /// <param name="parameter"><see cref="MessageArgs"/> argument
        /// for the <see cref="Message.Sent"/> event.</param>
        protected override void Invoke(object parameter)
        {
            dynamic args = parameter;
            if (args != null)
            {
                Content = args.Content;
                Window window = CreateView();
                // Only set the window's DataContext and handler for the
                // RequestClose event if the DataContext has not already been
                // assigned. We assume here that if a DataContext has been
                // assigned, it will have been done by a view models InjectInto
                // method, which also takes care of the close handler.
                if (window.DataContext == null)
                {
                    window.DataContext = this;
                    EventHandler closeHandler = null;
                    closeHandler = (sender, e) =>
                    {
                        Content.RequestCloseView -= closeHandler;
                        window.Close();
                        args.Respond();
                    };
                    Content.RequestCloseView += closeHandler;
                }
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
