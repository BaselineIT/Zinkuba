using System;

namespace Zinkuba.MailModule.API
{
    public interface IMessageReader<in TInMessageDescriptor> : IMessageReader
    {
        void Process(TInMessageDescriptor messageDescriptor);
    }

    public interface IMessageReader : IMessageProcessor
    {
        Type InMessageDescriptorType();
    }

    public interface IMessageConnector : IMessageReader, IMessageWriter
    {
    }

    public interface IMessageConnector<in TInMessageDescriptor, TOutMessageDescriptor> :
        IMessageReader<TInMessageDescriptor>, IMessageWriter<TOutMessageDescriptor>, IMessageConnector
    {
    }
}