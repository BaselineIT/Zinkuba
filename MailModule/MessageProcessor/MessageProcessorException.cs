using System;
using Zinkuba.MailModule.API;

namespace Zinkuba.MailModule.MessageProcessor
{
    public class MessageProcessorException : Exception
    {
        public MessageProcessorException(string message)
            : base(message)
        {
        }

        public MessageProcessorStatus Status { get; set; }
    }
}