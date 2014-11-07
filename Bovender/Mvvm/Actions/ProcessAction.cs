using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Bovender.Mvvm.Messaging;
using Bovender.Mvvm.Views;

namespace Bovender.Mvvm.Actions
{
    /// <summary>
    /// Invokes a process view.
    /// </summary>
    class ProcessAction : MessageActionBase
    {
        #region Overrides

        /// <summary>
        /// Injects the message <see cref="Content"/> into a newly created
        /// <see cref="ProcessView"/> and returns the view.
        /// </summary>
        /// <returns>Instance of <see cref="ProcessView"/> that is data bound
        /// to the current message Content.</returns>
        protected override System.Windows.Window CreateView()
        {
            if (Content is ProcessMessageContent)
            {
                return Content.InjectInto<ProcessView>();
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
