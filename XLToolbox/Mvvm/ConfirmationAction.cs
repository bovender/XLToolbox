using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Interactivity;
using XLToolbox.Core.Mvvm;

namespace XLToolbox.Mvvm
{
    public class ConfirmationAction : TriggerAction<FrameworkElement>
    {
        public string Caption { get; set; }
        public string Message { get; set; }
        public MessageContent Content { get; private set; }

        protected override void Invoke(object parameter)
        {
            MessageArgs<MessageContent> args = parameter as MessageArgs<MessageContent>;
            if (args != null)
            {
                Content = args.Content;
                WindowConfirmation window = new WindowConfirmation();
                window.DataContext = this;
                EventHandler closeHandler = null;
                closeHandler = (sender, e) =>
                {
                    window.Closed -= closeHandler;
                    args.Respond();
                };
                window.Closed += closeHandler;
                window.Show();
            }
 
        }
    }
}
