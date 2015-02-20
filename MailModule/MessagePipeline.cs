using System;
using System.Collections.Generic;
using log4net;
using Zinkuba.MailModule.API;
using Zinkuba.MailModule.MessageDescriptor;
using Zinkuba.MailModule.MessageProcessor;

namespace Zinkuba.MailModule
{

    public class MessagePipeline
    {
        private readonly List<IMessageProcessor> _messageProcessors;
        private static readonly ILog Logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        private MessageProcessorStatus _state;
        public int TotalMails { get; private set; }

        public MessageProcessorStatus State
        {
            get { return _state; }
            private set { _state = value; OnStateChanged(); }
        }

        private int _exportedMails;
        private int _failedMails;
        private int _ignoredMails;

        public int ExportedMails
        {
            get { return _exportedMails; }
            private set { _exportedMails = value; OnExportedMail(); }
        }

        public int FailedMails
        {
            get { return _failedMails; }
            private set { _failedMails = value; OnFailedMail(); }
        }

        public int IgnoredMails
        {
            get { return _ignoredMails; }
            private set { _ignoredMails = value; OnIgnoredMail(); }
        }

        public event EventHandler ExportedMail;

        protected virtual void OnExportedMail()
        {
            EventHandler handler = ExportedMail;
            if (handler != null) handler(this, EventArgs.Empty);
        }

        public event EventHandler FailedMail;

        protected virtual void OnFailedMail()
        {
            EventHandler handler = FailedMail;
            if (handler != null) handler(this, EventArgs.Empty);
        }

        public event EventHandler IgnoredMail;

        protected virtual void OnIgnoredMail()
        {
            EventHandler handler = IgnoredMail;
            if (handler != null) handler(this, EventArgs.Empty);
        }

        public event EventHandler StateChanged;

        protected virtual void OnStateChanged()
        {
            EventHandler handler = StateChanged;
            if (handler != null) handler(this, EventArgs.Empty);
        }

        public MessagePipeline(List<IMessageProcessor> messageProcessors)
        {
            _messageProcessors = messageProcessors;
            IMessageWriter previousWriter = null;
            foreach (var messageProcessor in messageProcessors)
            {
                messageProcessor.FailedMessage += (sender, args) => MessageFailed(messageProcessor.FailedMessageCount);
                messageProcessor.IgnoredMessage += (sender, args) => MessageIgnored(messageProcessor.IgnoredMessageCount);
                messageProcessor.StatusChanged += MessageProcessorOnStatusChanged;
                //var messageWriter = messageProcessor as IMessageWriter;
                if (previousWriter != null)
                {
                    try
                    {
                        ((dynamic)previousWriter).NextReader = (dynamic)messageProcessor;
                    }
                    catch (Exception e)
                    {
                        Logger.Error("Failed to connect up the pipeline : " + e.Message);
                        throw;
                    }
                }
                previousWriter = messageProcessor as IMessageWriter;
            }
            // Listen on success messge of last item
            messageProcessors[messageProcessors.Count - 1].SucceededMessage += (sender, args) => MessageExported(messageProcessors[messageProcessors.Count - 1].SucceededMessageCount);
            // listen on Total messages of first item
            ((IMessageSource)messageProcessors[0]).TotalMessagesChanged += OnTotalMessagesChanged;

            TotalMails = 0;
            ExportedMails = 0;
            State = MessageProcessorStatus.Idle;
        }

        private void OnTotalMessagesChanged(object sender, EventArgs eventArgs)
        {
            var messageSource = sender as IMessageSource;
            if (messageSource != null)
            {
                TotalMails = messageSource.TotalMessages;
            }
        }

        private void MessageProcessorOnStatusChanged(object sender, EventArgs eventArgs)
        {
            var processor = sender as IMessageProcessor;
            if (processor != null)
            {
                if (!Failed)
                {
                    State = processor.Status;
                }
            }
        }

        private void MessageExported(int exportedMails)
        {
            ExportedMails = exportedMails;
        }

        private void MessageIgnored(int ignoredMails)
        {
            IgnoredMails = ignoredMails;
        }

        private void MessageFailed(int failedMails)
        {
            FailedMails = failedMails;
        }

        public bool Running { get { return State == MessageProcessorStatus.Started || State == MessageProcessorStatus.Initialising; } }
        public bool Failed { get { return State == MessageProcessorStatus.AuthFailure || State == MessageProcessorStatus.ConnectionError || State == MessageProcessorStatus.UnknownError; } }
        public bool TestOnly { get; set; }

        public void Start()
        {
            State = MessageProcessorStatus.Idle;
            try
            {
                _messageProcessors[0].Initialise();
            }
            catch (MessageProcessorException ex)
            {
                State = ex.Status;
                throw ex;
            }
            catch (Exception ex)
            {
                Logger.Error("Failed to initialise service", ex);
                State = MessageProcessorStatus.UnknownError;
                throw ex;
            }
            ((IMessageSource)_messageProcessors[0]).Start();
        }

        public void Close(MessageProcessorStatus state = MessageProcessorStatus.Completed)
        {
            if (State == MessageProcessorStatus.Started || State == MessageProcessorStatus.Initialising)
            {
                _messageProcessors[0].Close();
                State = state;
            }
        }

        public void Cancel()
        {
            Close(MessageProcessorStatus.Cancelled);
        }

    }
}
