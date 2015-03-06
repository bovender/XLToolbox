/* Message.cs
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
using System.Linq;
using System.Text;

namespace Bovender.Mvvm.Messaging
{
    /// <summary>
    /// Conveys a message from a view model to a consumer (typically, a view)
    /// that has subscribed to the Sent event.
    /// </summary>
    /// <typeparam name="T">Type of the content of the message (must be a
    /// <see cref="MessageContent"/> object or a descendant.</typeparam>
    public class Message<T> : IMessage<T> where T : MessageContent, new() 
    {
        #region IViewModelMessage interface

        /// <summary>
        /// Consumers of the view model subscribe to this event if they want
        /// to listen for the message.
        /// </summary>
        public event EventHandler<MessageArgs<T>> Sent;

        #endregion

        #region Protected methods

        /// <summary>
        /// Calling this method will raise the Sent event with a message content
        /// and a callback method that can be used by the View to send a return signal.
        /// </summary>
        /// <param name="messageContent">Content of the message.</param>
        /// <param name="respond">Callback method that accepts a parameter of same type
        /// as <paramref name="messageContent"/>.</param>
        public virtual void Send(
            T messageContent,
            Action<T> respond)
        {
            if (Sent != null)
            {
                Sent(this,
                    new MessageArgs<T>(
                        messageContent,
                        () =>
                        { 
                            if (respond != null) respond(messageContent);
                        }
                    )
                );
            };
        }

        /// <summary>
        /// Raises the Sent event with a <see cref="MessageArgs"/> instance that
        /// encapsulates the <paramref name="messageContent"/>
        /// </summary>
        /// <param name="messageContent">Derivate of MessageContent that defines the
        /// message content.</param>
        public virtual void Send(T messageContent)
        {
            Send(messageContent, null);
        }

        /// <summary>
        /// Sends a simple message that does not need responding to.
        /// </summary>
        public virtual void Send()
        {
            Send(new T(), null);
        }

        #endregion
    }
}
