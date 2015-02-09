using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Zinkuba.MailModule.API
{
    public interface IMessageSource : IMessageProcessor
    {
        int TotalMessages { get; }
        void Start();
        bool TestOnly { get; set; }

        event EventHandler TotalMessagesChanged;
    }
}
