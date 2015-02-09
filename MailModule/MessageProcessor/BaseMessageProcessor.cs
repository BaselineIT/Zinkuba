using System;
using Zinkuba.MailModule.API;

namespace Zinkuba.MailModule.MessageProcessor
{
    public abstract class BaseMessageProcessor : IMessageProcessor
    {
        private int _processedMessageCount;
        private int _failedMessageCount;
        private int _succeededMessageCount;
        private int _ignoredMessageCount;
        private MessageProcessorStatus _status;
        public abstract void Initialise();
        public abstract void Close();
        public bool Closed { get; protected set; }
        private string _name;

        public string Name
        {
            get { return _name; }
            set { _name = value; }
        }

        protected BaseMessageProcessor()
        {
            Status = MessageProcessorStatus.Idle;
        }

        public int ProcessedMessageCount
        {
            get { return _processedMessageCount; }
            protected set { _processedMessageCount = value; OnProcessedMessage(); }
        }

        public int FailedMessageCount
        {
            get { return _failedMessageCount; }
            protected set { _failedMessageCount = value; OnFailedMessage(); }
        }

        public int SucceededMessageCount
        {
            get { return _succeededMessageCount; }
            protected set { _succeededMessageCount = value; OnSucceededMessage(); }
        }

        public int IgnoredMessageCount
        {
            get { return _ignoredMessageCount; }
            protected set { _ignoredMessageCount = value; OnIgnoredMessage(); }
        }

        public MessageProcessorStatus Status
        {
            get { return _status; }
            protected set { _status = value; OnStatusChanged(); }
        }

        public event EventHandler ProcessedMessage;

        protected virtual void OnProcessedMessage()
        {
            EventHandler handler = ProcessedMessage;
            if (handler != null) handler(this, EventArgs.Empty);
        }

        public event EventHandler FailedMessage;

        protected virtual void OnFailedMessage()
        {
            EventHandler handler = FailedMessage;
            if (handler != null) handler(this, EventArgs.Empty);
        }

        public event EventHandler SucceededMessage;

        protected virtual void OnSucceededMessage()
        {
            EventHandler handler = SucceededMessage;
            if (handler != null) handler(this, EventArgs.Empty);
        }

        public event EventHandler IgnoredMessage;

        protected virtual void OnIgnoredMessage()
        {
            EventHandler handler = IgnoredMessage;
            if (handler != null) handler(this, EventArgs.Empty);
        }

        public event EventHandler StatusChanged;

        protected virtual void OnStatusChanged()
        {
            EventHandler handler = StatusChanged;
            if (handler != null) handler(this, EventArgs.Empty);
        }
        public bool Running { get { return Status == MessageProcessorStatus.Started || Status == MessageProcessorStatus.Initialising; } }
        public bool Failed { get { return Status == MessageProcessorStatus.AuthFailure || Status == MessageProcessorStatus.ConnectionError || Status == MessageProcessorStatus.UnknownError; } }
    }
}
