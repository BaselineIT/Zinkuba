using System;
using Zinkuba.App.Annotations;
using Zinkuba.App.Mailbox;

namespace Zinkuba.App.MailAccount
{
    public interface IMailAccount
    {
        event EventHandler<IMailbox> AddedMailbox;
        event EventHandler<IMailbox> RemovedMailbox;

        ReturnTypeCollection<IMailbox> Mailboxes { get; }
    }
}
