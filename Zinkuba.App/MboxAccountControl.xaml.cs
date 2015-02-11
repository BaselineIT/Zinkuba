using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;
using Zinkuba.App.Annotations;
using Zinkuba.App.MailAccount;
using Zinkuba.App.Mailbox;

namespace Zinkuba.App
{
    /// <summary>
    /// Interaction logic for MboxAccountControl.xaml
    /// </summary>
    public partial class MboxAccountControl : UserControl, IMailAccountControl
    {
        public IMailAccount Account { get; private set; }
        public MBoxAccountDataContext _dataContext;
        public Action<IMailAccount> RemoveAccountFunction;

        public MboxAccountControl(MboxAccount account)
        {
            InitializeComponent();
            Account = account;
            account.Mailboxes.CollectionChanged += MailboxesOnCollectionChanged;
            _dataContext = new MBoxAccountDataContext(account, RemoveMailboxFunction);
            DataContext = _dataContext;
        }

        private void MailboxesOnCollectionChanged(object sender, NotifyCollectionChangedEventArgs notifyCollectionChangedEventArgs)
        {
            if (Account.Mailboxes.Count == 0)
            {
                RemoveItem();
            }
        }

        public IMailAccount MirrorSource { get { return Account; } }

        private void RemoveItem()
        {
            if (RemoveAccountFunction != null)
            {
                RemoveAccountFunction(Account);
            }
        }

        private void RemoveMailboxFunction(IMailbox o)
        {
            var item = o as MboxMailbox;
            if (item != null)
            {
                if (Account.Mailboxes.Contains<MboxMailbox>(item))
                {
                    Account.Mailboxes.Remove<MboxMailbox>(item);
                }
            }
        }

        public List<MboxMailbox> PendingMailboxes()
        {
            var pending = new List<MboxMailbox>();
            foreach (var mailboxControl in _dataContext.Mailboxes)
            {
                if ((mailboxControl.Mailbox.Exporter == null || mailboxControl.Mailbox.Exporter.Failed) && mailboxControl.Validate())
                {
                    pending.Add(mailboxControl.Mailbox);
                }
            }
            return pending;
        }
    }

    public class MBoxAccountDataContext : INotifyPropertyChanged
    {
        private readonly MboxAccount _account;
        public MboxMailboxList Mailboxes { get; set; }

        public event PropertyChangedEventHandler PropertyChanged;

        public MBoxAccountDataContext(MboxAccount account, Action<MboxMailbox> removeMailboxAction)
        {
            _account = account;
            Mailboxes = new MboxMailboxList(removeMailboxAction, Dispatcher.CurrentDispatcher);
            Mailboxes.SetList(account.Mailboxes.UnderlyingCollection);
            //Mailboxes.CollectionChanged += (sender, args) => OnPropertyChanged("Mailboxes");
        }

        [NotifyPropertyChangedInvocator]
        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChangedEventHandler handler = PropertyChanged;
            if (handler != null) handler(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
