using Bovender.Mvvm.Messaging;
using System;

namespace Bovender.Mvvm.Actions
{
    /// <summary>
    /// MVVM action that queries the user for a folder.
    /// </summary>
    /// <remarks>
    /// To be used with MVVM messages that carry a <see cref="StringMessageContent"/>.
    /// </remarks>
    public class ChooseFolderAction : MessageActionBase
    {
        #region Public properties

        public string Description { get; set; }

        #endregion

        #region Overrides

        protected override void Invoke(object parameter)
        {
            MessageArgs<StringMessageContent> args = parameter as MessageArgs<StringMessageContent>;
            System.Windows.Forms.FolderBrowserDialog dlg = new System.Windows.Forms.FolderBrowserDialog();
            dlg.SelectedPath = args.Content.Value;
            dlg.ShowNewFolderButton = true;
            dlg.Description = Description;
            dlg.ShowDialog();
            args.Content.Confirmed = !string.IsNullOrEmpty(dlg.SelectedPath);
            if (args.Content.Confirmed)
            {
                args.Content.Value = dlg.SelectedPath;
            };
            args.Respond();
        }

        protected override System.Windows.Window CreateView()
        {
            throw new InvalidOperationException("The ChooseFolderAction does not create an MVVM WPF - this method must not be called.");
        }

        #endregion
    }
}
