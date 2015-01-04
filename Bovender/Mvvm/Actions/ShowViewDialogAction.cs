/* ShowViewDialogAction.cs
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
using System.Windows;

namespace Bovender.Mvvm.Actions
{
    /// <summary>
    /// Injects a view with a view model that is referenced in a ViewModelMessageContent,
    /// and shows the view modally.
    /// </summary>
    public class ShowViewDialogAction : ShowViewAction
    {
        protected override void ShowView(Window view)
        {
            view.ShowDialog();
        }
    }
}
