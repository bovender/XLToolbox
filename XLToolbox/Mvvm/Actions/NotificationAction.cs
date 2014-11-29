using System;
using Bovender.Mvvm.Actions;
using XLToolbox.Mvvm.Views;

namespace XLToolbox.Mvvm.Actions
{
    class NotificationAction : MessageActionBase
    {
        protected override System.Windows.Window CreateView()
        {
            return new NotificationView();
        }
    }
}
