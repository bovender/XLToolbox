/* MessageContent.cs
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
using Bovender.Mvvm.ViewModels;

namespace Bovender.Mvvm.Messaging
{
    /// <summary>
    /// Simple object that encapsulates a boolean value; to be used
    /// in MVVM interaction with <see cref="MessageArgs"/>.
    /// </summary>
    public class MessageContent : ViewModelBase
    {
        #region Public properties

        public bool Confirmed { get; set; }

        #endregion

        #region Commands

        /// <summary>
        /// Sets the <see cref="Confirmed"/> property to true and
        /// triggers a <see cref="RequestCloseView"/> event. 
        /// </summary>
        public DelegatingCommand ConfirmCommand
        {
            get
            {
                if (_confirmCommand == null)
                {
                    _confirmCommand = new DelegatingCommand(
                        (param) => { DoConfirm(); },
                        (param) => { return CanConfirm(); }
                        );
                };
                return _confirmCommand;
            }
        }

        #endregion

        #region Constructors

        /// <summary>
        /// Creates a new, empty message content.
        /// </summary>
        public MessageContent() : base() { }

        #endregion

        #region Protected methods

        /// <summary>
        /// Executes the confirmation logic: sets <see cref="Confirmed"/> to True
        /// and calls <see cref="DoCloseView()"/> to issue a RequestCloseView
        /// message.
        /// </summary>
        protected virtual void DoConfirm()
        {
            Confirmed = true;
            DoCloseView();
        }

        /// <summary>
        /// Determines whether the ConfirmCommand can be executed.
        /// </summary>
        /// <returns>True if the ConfirmCommand can be executed.</returns>
        protected virtual bool CanConfirm()
        {
            return true;
        }

        #endregion

        #region Private properties

        private DelegatingCommand _confirmCommand;

        #endregion

        public override object RevealModelObject()
        {
            return null;
        }
    }
}
