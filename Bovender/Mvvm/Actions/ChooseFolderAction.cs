using System;
using System.Windows.Forms;
using Bovender.Mvvm.Messaging;

namespace Bovender.Mvvm.Actions
{
    /// <summary>
    /// MVVM action that queries the user for a folder.
    /// </summary>
    /// <remarks>
    /// To be used with MVVM messages that carry a <see cref="StringMessageContent"/>.
    /// </remarks>
    public class ChooseFolderAction : FileFolderActionBase
    {
        protected override string GetDialogResult(string defaultString)
        {
            FolderBrowserDialog dlg = new FolderBrowserDialog();
            dlg.SelectedPath = defaultString;
            dlg.ShowNewFolderButton = true;
            dlg.Description = Description;
            if (dlg.ShowDialog() == DialogResult.OK)
            {
                return dlg.SelectedPath;
            }
            else
            {
                return String.Empty;
            }
        }
    }
}
