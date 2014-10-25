using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using XLToolbox.Views;

namespace XLToolbox.Mvvm
{
    class StringMessageAction : ConfirmationAction
    {
        protected override System.Windows.Window CreateView()
        {
            return new StringMessageContentView();
        }
    }
}
