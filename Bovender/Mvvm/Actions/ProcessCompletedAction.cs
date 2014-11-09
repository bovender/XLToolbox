using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using Bovender.Mvvm.Messaging;

namespace Bovender.Mvvm.Actions
{
    /// <summary>
    /// Abstract WPF action that invokes different views depending on the status
    /// of a completed process.
    /// </summary>
    public abstract class ProcessCompletedAction : MessageActionBase
    {
        #region Abstract methods

        protected abstract Window CreateSuccessWindow();
        protected abstract Window CreateFailureWindow();
        protected abstract Window CreateCancelledWindow();

        #endregion

        #region Overrides

        protected override Window CreateView()
        {
            ProcessMessageContent content = Content as ProcessMessageContent;
            if (Content is ProcessMessageContent)
            {
                if (content.WasCancelled)
                {
                    return content.InjectInto(CreateCancelledWindow());
                }
                else if (content.WasSuccessful)
                {
                    return content.InjectInto(CreateSuccessWindow());
                }
                else
                {
                    return content.InjectInto(CreateFailureWindow());
                }
            }
            else
            {
                throw new ArgumentException(
                    "This message action must be used for Messages with ProcessMessageContent only.");
            }
        }

        #endregion
    }
}
