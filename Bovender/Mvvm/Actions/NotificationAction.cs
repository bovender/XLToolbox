/* NotificationAction.cs
 * part of Daniel's XL Toolbox NG
 * 
 * Copyright 2014-2015 Daniel Kraus
 * 
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 * 
 *     http://www.apache.org/licenses/LICENSE-2.0
 * 
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
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
