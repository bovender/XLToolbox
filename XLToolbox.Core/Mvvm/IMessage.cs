using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace XLToolbox.Core.Mvvm
{
    /// <summary>
    /// Defines a Sent event that consumers of a view model can
    /// subscribe to in order to listen to the view model's message.
    /// </summary>
    public interface IMessage
    {
        event EventHandler<MessageArgs> Sent;
    }
}
