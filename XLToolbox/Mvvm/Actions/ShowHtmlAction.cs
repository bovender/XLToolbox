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
        protected override Window CreateView()
        {
            StringMessageContent smc = Content as StringMessageContent;
            if (smc != null)
            {
                HtmlFileViewModel vm = new HtmlFileViewModel(smc.Value);
                Window view = vm.InjectInto<HtmlFileView>();
                return view;
            }
            else
            {
                throw new InvalidOperationException(
                    "Message content for ShowHtmlAction must be a StringMessageContent");
            }
        }
    }
}
