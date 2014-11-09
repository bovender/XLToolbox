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
            if (!string.IsNullOrEmpty(HtmlResource))
            {
                HtmlFileViewModel vm = new HtmlFileViewModel(HtmlResource);
                vm.Caption = Caption;
                Window view = vm.InjectInto<HtmlFileView>();
                return view;
            }
            else
            {
                throw new ArgumentNullException(
                    "Must assign HtmlResource field in ShowHtmlAction tag.");
            }
        }
    }
}
