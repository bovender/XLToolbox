/* ShowViewAction.cs
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
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using Bovender.Mvvm.Messaging;

namespace Bovender.Mvvm.Actions
{
    /// <summary>
    /// Injects a view with a view model that is referenced in a ViewModelMessageContent,
    /// and shows the view non-modally.
    /// </summary>
    public class ShowViewAction : MessageActionBase
    {
        #region Public properties

        public string Assembly { get; set; }
        public string View { get; set; }

        #endregion

        #region Overrides

        protected override Window CreateView()
        {
            object obj = Activator.CreateInstance(Assembly, View).Unwrap();
            Window view = obj as Window;
            if (view != null)
            {
                ViewModelMessageContent content = Content as ViewModelMessageContent;
                content.ViewModel.InjectInto(view);
                return view;
            }
            else
            {
                throw new ArgumentException(String.Format(
                    "Class name '{0}' in assembly '{1}' is not derived from Window.",
                    Assembly, View));
            }
        }

        protected override void ShowView(Window view)
        {
            view.Show();
        }

        #endregion
    }
}
