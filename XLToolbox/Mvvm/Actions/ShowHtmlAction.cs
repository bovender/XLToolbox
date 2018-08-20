/* ShowHtmlAction.cs
 * part of Daniel's XL Toolbox NG
 * 
 * Copyright 2014-2018 Daniel Kraus
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
using Bovender.Mvvm.ViewModels;
using Bovender.HtmlFiles;
using XLToolbox.Mvvm.Views;

namespace XLToolbox.Mvvm.Actions
{
    class ShowHtmlAction : StringMessageAction
    {
        public string HtmlResource { get; set; }

        protected override Window CreateView()
        {
            return new HtmlFileView();
        }

        protected override ViewModelBase GetDataContext(MessageContent messageContent)
        {
            if (!string.IsNullOrEmpty(HtmlResource))
            {
                HtmlFileViewModel vm = new HtmlFileViewModel(HtmlResource);
                vm.Caption = Caption;
                return vm;
            }
            else
            {
                throw new ArgumentNullException(
                    "Must assign HtmlResource field in ShowHtmlAction tag.");
            }
        }
    }
}
