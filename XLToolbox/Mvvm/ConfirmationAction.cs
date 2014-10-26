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
            dynamic args = parameter;
            if (args != null)
            {
                Content = args.Content;
                Window window = CreateView();
                window.DataContext = this;
                EventHandler closeHandler = null;
                closeHandler = (sender, e) =>
                {
                    window.Closed -= closeHandler;
                    args.Respond();
                };
                window.Closed += closeHandler;
                window.ShowDialog();
            }
        }

        /// <summary>
        /// Returns a view that can bind to expected message contents.
        /// </summary>
        /// <returns>Descendant of Window.</returns>
        protected virtual Window CreateView()
        {
            return new WindowConfirmation();
        }
    }
}
