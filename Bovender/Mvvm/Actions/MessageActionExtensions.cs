/* MessageActionExtensions.cs
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
            // Call Invoke with dummy message args and message content.
            action.Invoke(
                new Messaging.MessageArgs<Messaging.MessageContent>(
                    new Messaging.MessageContent(), null
                )
            );
        }
    }
}