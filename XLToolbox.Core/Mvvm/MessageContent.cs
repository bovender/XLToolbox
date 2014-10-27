namespace XLToolbox.Core.Mvvm
{
    /// <summary>
    /// Simple object that encapsulates a boolean value; to be used
    /// in MVVM interaction with <see cref="MessageArgs"/>.
    /// </summary>
    public class MessageContent : ViewModelBase
    {
        #region Private properties

        private DelegatingCommand _confirmCommand;

        #endregion

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

        #region Protected methods

        protected virtual void DoConfirm()
        {
            Confirmed = true;
            DoCloseView();
        }

        protected virtual bool CanConfirm()
        {
            return true;
        }

        #endregion
    }
}
