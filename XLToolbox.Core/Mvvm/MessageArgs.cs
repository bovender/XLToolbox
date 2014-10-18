using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace XLToolbox.Core.Mvvm
{
    public class MessageArgs : EventArgs
    {
        public MessageContent Content { get; set; }
        public Action Respond { get; set; }

        public MessageArgs(MessageContent content, Action respond)
        {
            Content = content;
            Respond = respond;
        }
    }
}
