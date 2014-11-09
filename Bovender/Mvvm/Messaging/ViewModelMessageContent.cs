using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Bovender.Mvvm.ViewModels;

namespace Bovender.Mvvm.Messaging
{
    /// <summary>
    /// Message content that holds a reference to a view model.
    /// </summary>
    public class ViewModelMessageContent : MessageContent
    {
        #region Public properties

        public ViewModelBase ViewModel { get; set; }

        #endregion

        #region Constructors

        public ViewModelMessageContent() : base() { }

        /// <summary>
        /// Instantiates the ViewModelMessageContent with a given view model.
        /// </summary>
        /// <param name="viewModel">Descendant of <see cref="ViewModelBase"/></param>
        public ViewModelMessageContent(ViewModelBase viewModel)
            : this()
        {
            ViewModel = viewModel;
        }

        #endregion
    }
}
