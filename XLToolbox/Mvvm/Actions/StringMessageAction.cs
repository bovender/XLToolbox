using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Bovender.Mvvm.Actions;
using XLToolbox.Mvvm.Views;

namespace XLToolbox.Mvvm.Actions
{
    class StringMessageAction : MessageActionBase
    {
        protected override System.Windows.Window CreateView()
        {
            return new StringMessageContentView();
        }
    }
}
