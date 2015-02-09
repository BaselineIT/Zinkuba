using System;
using System.Collections.Generic;
using Zinkuba.MailModule.API;

namespace Zinkuba.MailModule.MessageProcessor
{
    public interface IMessageDestination : IMessageProcessor
    {
        List<String> ImportedIds { get; }
    }
}