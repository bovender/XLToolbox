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
    /// Invokes a process view and injects the ProcessMessageContent
    /// as a view model into it.
    /// </summary>
    /// <remarks>
    /// This action cannot inject itself into the view because actions
    /// are not view models by Bovender's definition. To enable a view
    /// that invoke this action to set the strings itself, the Caption
    /// and Message properties of the MessageActionBase parent class
    /// and a CancelButtonText are written to the message content
    /// object (if they are not null or empty strings) so that they
    /// are available in the newly created view that binds the message
    /// content as its view model.
    /// </remarks>
    public class ProcessAction : ShowViewAction
    {
        #region Public properties

        public string CancelButtonText { get; set; }

        #endregion

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
                if (!string.IsNullOrEmpty(Caption)) pcm.Caption = Caption;
                if (!string.IsNullOrEmpty(Message)) pcm.Message = Message;
                if (!string.IsNullOrEmpty(CancelButtonText)) pcm.CancelButtonText = CancelButtonText;
                Window view;
                // Attempt to create a view from the Assembly and View
                // parameters. If this fails, create a generic ProcessView.
                try
                {
                    view = base.CreateView();
                }
                catch
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
