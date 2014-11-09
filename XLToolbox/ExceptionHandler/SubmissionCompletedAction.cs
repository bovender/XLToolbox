using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using Bovender.Mvvm.Actions;

namespace XLToolbox.ExceptionHandler
{
    /// <summary>
    /// WPF action that is invoked when the exception report submission
    /// process is completed.
    /// </summary>
    class SubmissionCompletedAction : ProcessCompletedAction
    {
        protected override Window CreateSuccessWindow()
        {
            return Content.InjectInto<SubmissionSuccessView>();
        }

        protected override Window CreateFailureWindow()
        {
            throw new NotImplementedException();
        }

        protected override Window CreateCancelledWindow()
        {
            throw new NotImplementedException();
        }
    }
}
