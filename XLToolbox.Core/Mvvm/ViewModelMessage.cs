using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace XLToolbox.Core.Mvvm
{
    /// <summary>
    /// Conveys a message from a view model to a consumer (typically, a view)
    /// that has subscribed to the Sent event.
    /// </summary>
    /// <typeparam name="T">Type of the content of the message.</typeparam>
    class ViewModelMessage<T> : IViewModelMessage
    {
        #region IViewModelMessage interface

        /// <summary>
        /// Consumers of the view model subscribe to this event if they want
        /// to listen for the message.
        /// </summary>
        public event EventHandler<ViewModelMessageArgs> Sent;

        #endregion

        #region Protected methods

        /// <summary>
        /// Calling this method will raise the Sent event with a message content
        /// and a callback method that can be used by the View to send a return signal.
        /// </summary>
        /// <param name="messageContent">Content of the message.</param>
        /// <param name="respond">Callback method that accepts a parameter of same type as <paramref name="messageContent"/>.</param>
        protected virtual void Send(T messageContent, Action<T> respond)
        {
            if (Sent != null)
            {
                Sent(this,
                    new ViewModelMessageArgs(
                        messageContent,
                        () => respond(messageContent)
                    )
                );
            };
        }

        #endregion
    }
}
