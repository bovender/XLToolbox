using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using Bovender.Mvvm.Actions;
using Bovender.Mvvm.Messaging;
using System.Windows.Forms;

namespace XLToolbox.Mvvm.Actions
{
    /// <summary>
    /// WPF action that displays a folder picker dialog and returns the chosen folder
    /// in the string Value field of a <see cref="StringMessageContent"/>.
    /// </summary>
    class ChooseFolderAction : MessageActionBase
    {
        #region Overrides

        protected override void Invoke(object parameter)
        {
            MessageArgs<StringMessageContent> args = parameter as MessageArgs<StringMessageContent>;
            if (args != null)
            {
                args.Content.Confirmed = false;
                FolderBrowserDialog dlg = new FolderBrowserDialog();
                dlg.SelectedPath = args.Content.Value;
                dlg.ShowNewFolderButton = true;
                dlg.ShowDialog();
                if (!string.IsNullOrEmpty(dlg.SelectedPath))
                {
                    args.Content.Value = dlg.SelectedPath;
                    args.Content.Confirmed = true;
                    args.Respond();
                }
            }
            else
            {
                throw new InvalidOperationException("Expected to receive Message<StringMessageContent> as parameter.");
            }
        }

        /// <summary>
        /// Dummy implementation of the abstract method in the parent class.
        /// Will not be called because this class also overrides <see cref="Invoke"/>.
        /// </summary>
        /// <exception cref="InvalidOperationException">If this method is called (which
        /// it shouldn't, by design).</exception>
        /// <returns>Nothing.</returns>
        protected override Window CreateView()
        {
            throw new InvalidOperationException("This method should never be invoked in this class.");
        }

        #endregion
    }
}
