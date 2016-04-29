using System;
using Zinkuba.App.Annotations;
using Zinkuba.App.Mailbox;

namespace Zinkuba.App.MailAccount
{
    public interface IMailAccount
    {
        event EventHandler<IMailbox> AddedMailbox;
        event EventHandler<IMailbox> RemovedMailbox;
        DateTime StartDate { get; set; }
        DateTime EndDate { get; set; }

        ReturnTypeCollection<IMailbox> Mailboxes { get; }
        String LimitFolder { get; set; }
    }
}
