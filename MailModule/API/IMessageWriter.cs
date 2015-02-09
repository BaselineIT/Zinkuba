using System;

namespace Zinkuba.MailModule.API
{
    public interface IMessageWriter<TOutMessageDescriptor> : IMessageWriter
    {
        IMessageReader<TOutMessageDescriptor> NextReader { get; set; }
    }

    public interface IMessageWriter : IMessageProcessor
    {
        Type OutMessageDescriptorType();
    }
}