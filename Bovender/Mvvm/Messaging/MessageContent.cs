using Bovender.Mvvm.ViewModels;

namespace Bovender.Mvvm.Messaging
{
    /// <summary>
    /// Simple object that encapsulates a boolean value; to be used
    /// in MVVM interaction with <see cref="MessageArgs"/>.
    /// </summary>
    public class MessageContent : ViewModelBase
    {
        #region Public properties

        public bool Confirmed { get; set; }

        #endregion

        #region Commands

        /// <summary>
        /// Sets the <see cref="Confirmed"/> property to true and
        /// triggers a <see cref="RequestCloseView"/> event. 
        /// </summary>
        public DelegatingCommand ConfirmCommand
        {
            get
            {
                if (_confirmCommand == null)
                {
                    _confirmCommand = new DelegatingCommand(
                        (param) => { DoConfirm(); },
                        (param) => { return CanConfirm(); }
                        );
                };
                return _confirmCommand;
            }
        }

        #endregion

        #region Constructors

        /// <summary>
        /// Creates a new, empty message content.
        /// </summary>
        public MessageContent() : base() { }

        #endregion

        #region Protected methods

        /// <summary>
        /// Executes the confirmation logic: sets <see cref="Confirmed"/> to True
        /// and calls <see cref="DoCloseView()"/> to issue a RequestCloseView
        /// message.
        /// </summary>
        protected virtual void DoConfirm()
        {
            Confirmed = true;
            DoCloseView();
        }

        /// <summary>
        /// Determines whether the ConfirmCommand can be executed.
        /// </summary>
        /// <returns>True if the ConfirmCommand can be executed.</returns>
        protected virtual bool CanConfirm()
        {
            return true;
        }

        #endregion

        #region Private properties

        private DelegatingCommand _confirmCommand;

        #endregion

        public override object RevealModelObject()
        {
            return null;
        }
    }
}
