using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Bovender.Mvvm.Messaging
{
    public class MessageArgs<T> : EventArgs where T : MessageContent
    {
        public T Content { get; set; }
        public Action Respond { get; set; }

        public MessageArgs(T content, Action respond)
        {
            Content = content;
            Respond = respond;
        }
    }
}
