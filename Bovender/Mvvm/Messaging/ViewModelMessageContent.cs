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

        /// <remarks>
        /// If views and their actions (derived from <see cref="MessageActionBase"/>)
        /// need access to the view model that sent the message, a reference to the
        /// view model instance can be transmitted via the <see cref="ViewModel"/>
        /// property.
        /// </remarks>
        public ViewModelBase ViewModel { get; set; }

        #endregion

        #region Constructors

        public ViewModelMessageContent() : base() { }

        /// <summary>
        /// Creates a new instance and set the <see cref="ViewModel"/> property
        /// to the parameter.
        /// </summary>
        /// <param name="viewModel">Instance of a ViewModelBase or descendant
        /// (typically the originator of the message).</param>
        public ViewModelMessageContent(ViewModelBase viewModel)
            : this()
        {
            ViewModel = viewModel;
        }

        #endregion
    }
}
