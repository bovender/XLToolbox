/* ShortcutViewModelCollection.cs
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
using Bovender.Mvvm;
using Bovender.Mvvm.ViewModels;

namespace XLToolbox.Keyboard
{
    public class ShortcutViewModelCollection : ViewModelCollection<Shortcut, ShortcutViewModel>
    {
        public ShortcutViewModelCollection() : base(Manager.Default.Shortcuts) { }

        protected override ShortcutViewModel CreateViewModel(Shortcut model)
        {
            return new ShortcutViewModel(model);
        }
    }
}
