using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Zinkuba.App.Mailbox;

namespace Zinkuba.App.MailAccount
{
    public class ImapAccount : ServerMailAccount
    {
        private ObservableCollection<AuthenticatedMailbox> _mailboxes;
        private bool _useSsl;

        public bool UseSsl;

        public ImapAccount() : base()
        {
            _mailboxes = new ObservableCollection<AuthenticatedMailbox>();
            Mailboxes.UnderlyingCollection = _mailboxes;
            _mailboxes.CollectionChanged += MailboxesOnCollectionChanged;
        }

        public bool Gmail { get; set; }
    }
}
