using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Bovender.Mvvm.Actions;

namespace XLToolbox.ExceptionHandler
{
    class SubmissionProcessAction : MessageActionBase
    {
        protected override System.Windows.Window CreateView()
        {
            return Content.InjectInto<SubmissionProcessView>();
        }
    }
}
