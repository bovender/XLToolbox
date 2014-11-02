using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using Bovender.Mvvm.Messaging;

namespace Bovender.Mvvm.Actions
{
    /// <summary>
    /// Simple WPF action that closes a progress window.
    /// </summary>
    class ProcessCompletedAction : MessageActionBase
    {
        protected override void Invoke(object parameter)
        {
            MessageArgs<ProcessMessageContent> args = parameter as MessageArgs<ProcessMessageContent>;
            if (args != null)
            {
                args.Content.CloseViewCommand.Execute(null);
            }
            else
            {
                throw new ArgumentException(
                    "This message action must be used for Messages with ProcessMessageContent only.");
            }
        }

        protected override System.Windows.Window CreateView()
        {
            throw new InvalidOperationException(
                "This MessageAction derivative does not create new views.");
        }
    }
}
