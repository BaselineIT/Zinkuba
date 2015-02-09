using System;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using Zinkuba.App.Mailbox;

namespace Zinkuba.App.MailAccount
{
    public abstract class ServerMailAccount : IMailAccount
    {
        public String Server;
        public DateTime StartDate;
        public DateTime EndDate;

        protected ServerMailAccount()
        {
            Mailboxes = new ReturnTypeCollection<IMailbox>();
        }

        protected void MailboxesOnCollectionChanged(object sender, NotifyCollectionChangedEventArgs notifyCollectionChangedEventArgs)
        {
            if (notifyCollectionChangedEventArgs.NewItems != null)
            {
                foreach (var newItem in notifyCollectionChangedEventArgs.NewItems)
                {
                    OnAddedMailbox((IMailbox)newItem);
                }
            }
            if (notifyCollectionChangedEventArgs.OldItems != null)
            {
                foreach (var oldItem in notifyCollectionChangedEventArgs.OldItems)
                {
                    OnRemovedMailbox((IMailbox)oldItem);
                }
            }
        }

        public event EventHandler<IMailbox> AddedMailbox;
        public event EventHandler<IMailbox> RemovedMailbox;
        public ReturnTypeCollection<IMailbox> Mailboxes { get; private set; }

        protected virtual void OnRemovedMailbox(IMailbox e)
        {
            EventHandler<IMailbox> handler = RemovedMailbox;
            if (handler != null) handler(this, e);
        }

        protected virtual void OnAddedMailbox(IMailbox e)
        {
            EventHandler<IMailbox> handler = AddedMailbox;
            if (handler != null) handler(this, e);
        }
    }
}
