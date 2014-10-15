﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace XLToolbox.Core.Mvvm
{
    public class ViewModelMessageArgs : EventArgs
    {
        public ViewModelMessageConfirmation Content { get; set; }
        public Action Respond { get; set; }

        public ViewModelMessageArgs(ViewModelMessageConfirmation content, Action respond)
        {
            Content = content;
            Respond = respond;
        }
    }
}
