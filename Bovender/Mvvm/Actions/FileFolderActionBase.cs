using Bovender.Mvvm.Messaging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Bovender.Mvvm.Actions
{
    /// <summary>
    /// Abstract base class for the <see cref="ChooseFileSaveAction"/> and
    /// <see cref="ChooseFolderAction"/> classes.
    /// </summary>
    public abstract class FileFolderActionBase : MessageActionBase
    {
        #region Public properties

        public string Description { get; set; }

        #endregion

        #region Abstract methods

        /// <summary>
        /// Displays an dialog to choose file or folder names.
        /// </summary>
        /// <param name="defaultString">Indicates the default
        /// path and/or file name and/or extension.</param>
        /// <returns>Valid file name/path, or empty string if
        /// the dialog was cancelled.</returns>
        protected abstract string GetDialogResult(string defaultString);

        #endregion

        #region Protected properties

        protected StringMessageContent MessageContent { get; private set; }

        #endregion

        #region Overrides

        protected override void Invoke(object parameter)
        {
            MessageArgs<StringMessageContent> args = parameter as MessageArgs<StringMessageContent>;
            MessageContent = args.Content;
            string result = GetDialogResult(args.Content.Value);
            args.Content.Confirmed = !string.IsNullOrEmpty(result);
            if (args.Content.Confirmed)
            {
                args.Content.Value = result;
            };
            args.Respond();
        }

        protected override System.Windows.Window CreateView()
        {
            throw new InvalidOperationException(
                "This class does not create WPF views and this method should never be called.");
        }

        #endregion
    }
}
