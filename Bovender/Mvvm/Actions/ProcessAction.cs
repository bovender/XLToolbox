using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using Bovender.Mvvm.Messaging;
using Bovender.Mvvm.Views;

namespace Bovender.Mvvm.Actions
{
    /// <summary>
    /// Invokes a process view.
    /// </summary>
    public class ProcessAction : ShowViewAction
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
            ProcessMessageContent pcm = Content as ProcessMessageContent;
            if (pcm != null)
            {
                pcm.Caption = Caption;
                pcm.Message = Message;
                Window view;
                // Attempt to create a view from the Assembly and View
                // parameters. If this fails, create a generic ProcessView.
                try
                {
                    view = base.CreateView();
                }
                catch (Exception e)
                {
                    view = new ProcessView();
                }
                return Content.InjectInto(view);
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
