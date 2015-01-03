using System;
using System.Windows.Forms;
using System.IO;
using System.Text.RegularExpressions;

namespace Bovender.Mvvm.Actions
{
    /// <summary>
    /// Lets the user choose a file name for saving a file.
    /// </summary>
    public class ChooseFileSaveAction : FileDialogActionBase
    {
        protected override FileDialog GetDialog(string defaultString)
        {
            SaveFileDialog dlg = new SaveFileDialog();
            Regex regex = new Regex(@"[^*]*\*\..+");
            if (regex.IsMatch(defaultString))
            {
                dlg.Filter = Path.GetFileName(defaultString);
            }
            dlg.InitialDirectory = Bovender.FileHelpers.GetDirectoryName(defaultString);
            string fn = Path.GetFileNameWithoutExtension(defaultString);
            dlg.AddExtension = true;
            dlg.RestoreDirectory = true;
            dlg.SupportMultiDottedExtensions = true;
            dlg.ValidateNames = true;
            dlg.ShowHelp = false;
            dlg.OverwritePrompt = true;
            return dlg;
        }
    }
}
