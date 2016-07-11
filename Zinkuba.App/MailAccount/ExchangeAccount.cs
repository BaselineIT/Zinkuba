using System.Collections.ObjectModel;
using Zinkuba.App.Mailbox;

namespace Zinkuba.App.MailAccount
{
    public class ExchangeAccount : ServerMailAccount
    {
        private readonly ObservableCollection<AuthenticatedMailbox> _mailboxes;

        public ExchangeAccount()
            : base()
        {
            _mailboxes = new ObservableCollection<AuthenticatedMailbox>();
            Mailboxes.UnderlyingCollection = _mailboxes;
            _mailboxes.CollectionChanged += MailboxesOnCollectionChanged;
        }

        public bool IncludePublicFolders { get; set; }

    }
}
