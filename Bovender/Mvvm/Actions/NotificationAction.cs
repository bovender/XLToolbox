using System;
using System.Windows;
using Bovender.Mvvm.Views;

namespace Bovender.Mvvm.Actions
{
    /// <summary>
    /// Opens a generic WPF dialog window that displays a message and
    /// has a single OK button. The message string can include parameters.
    /// </summary>
    public class NotificationAction : MessageActionBase
    {
        #region Public (dependency) properties

        public string Param1
        {
            get { return (string)GetValue(Param1Property); }
            set
            {
                SetValue(Param1Property, value);
            }
        }

        public string Param2
        {
            get { return (string)GetValue(Param2Property); }
            set
            {
                SetValue(Param2Property, value);
            }
        }

        public string Param3
        {
            get { return (string)GetValue(Param3Property); }
            set
            {
                SetValue(Param3Property, value);
            }
        }

        public string OkButtonLabel
        {
            get { return (string)GetValue(OkButtonLabelProperty); }
            set
            {
                SetValue(OkButtonLabelProperty, value);
            }
        }

        /// <summary>
        /// Returns the <see cref="Message"/> string formatted with the
        /// three params.
        /// </summary>
        public string FormattedText
        {
            get
            {
                try
                {
                    return String.Format(Message, Param1, Param2, Param3);
                }
                catch
                {
                    return "*** No message text given! ***";
                }
            }
        }

        #endregion

        #region Declarations of dependency properties

        public static readonly DependencyProperty Param1Property = DependencyProperty.Register(
            "Param1", typeof(string), typeof(NotificationAction));

        public static readonly DependencyProperty Param2Property = DependencyProperty.Register(
            "Param2", typeof(string), typeof(NotificationAction));

        public static readonly DependencyProperty Param3Property = DependencyProperty.Register(
            "Param3", typeof(string), typeof(NotificationAction));

        public static readonly DependencyProperty OkButtonLabelProperty = DependencyProperty.Register(
            "OkButtonLabel", typeof(string), typeof(NotificationAction));

        #endregion

        #region Constructor

        public NotificationAction()
            : base()
        {
            OkButtonLabel = "OK";
        }

        #endregion

        #region Implementation of abstract base methods

        protected override Window CreateView()
        {
            return new NotificationView();
        }

        #endregion
    }
}
