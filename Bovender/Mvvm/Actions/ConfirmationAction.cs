using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using Bovender.Mvvm.Views;

namespace Bovender.Mvvm.Actions
{
    /// <summary>
    /// Provides an action to confirm or cancel a message.
    /// Adds a CancelButtonLabel to the NotificationAction
    /// base class and creates a ConfirmationView rather
    /// than a NotificationView.
    /// </summary>
    public class ConfirmationAction : NotificationAction
    {
        #region Dependency properties

        public string CancelButtonLabel
        {
            get { return (string)GetValue(CancelButtonLabelProperty); }
            set
            {
                SetValue(CancelButtonLabelProperty, value);
            }
        }

        #endregion

        #region Declarations of dependency properties

        public static readonly DependencyProperty CancelButtonLabelProperty = DependencyProperty.Register(
            "CancelButtonLabel", typeof(string), typeof(ConfirmationAction));

        #endregion

        #region Constructor

        public ConfirmationAction()
            : base()
        {
            CancelButtonLabel = "Cancel";
        }

        #endregion

        #region Overrides

        protected override System.Windows.Window CreateView()
        {
            return new ConfirmationView();
        }

        #endregion
    }
}
