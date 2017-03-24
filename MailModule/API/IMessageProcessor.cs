using System;
using System.Collections.Generic;

namespace Zinkuba.MailModule.API
{
    public enum MessageProcessorStatus
    {
        Idle, Initialising, Started, AuthFailure, Completed, UnknownError, Cancelled,
        ConnectionError,
        Initialised,
        SourceAuthFailure,
        DestinationAuthFailure,
    }

    public interface IMessageProcessor
    {
        int SucceededMessageCount { get; }
        int ProcessedMessageCount { get; }
        int FailedMessageCount { get; }
        int IgnoredMessageCount { get; }
        MessageProcessorStatus Status { get; }
        void Initialise(List<MailFolder> folderList);
        void Close();
        bool Closed { get; }
        bool Running { get; }
        bool Failed { get; }
        String Name { get; set; }

        event EventHandler ProcessedMessage;
        event EventHandler FailedMessage;
        event EventHandler SucceededMessage;
        event EventHandler IgnoredMessage;
        event EventHandler StatusChanged;
    }
}