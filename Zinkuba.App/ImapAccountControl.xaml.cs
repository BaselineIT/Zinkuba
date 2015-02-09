using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;
using log4net;
using Zinkuba.App.Annotations;
using Zinkuba.App.MailAccount;
using Zinkuba.App.Mailbox;

namespace Zinkuba.App
{
    /// <summary>
    /// Interaction logic for ImapAccount.xaml
    /// </summary>
    public partial class ImapAccountControl : UserControl, IMailAccountControl
    {
        private static readonly ILog Logger = LogManager.GetLogger(typeof (ImapAccountControl));
        private ImapAccountDataContext _dataContext;
        public Action<IMailAccount> RemoveAccountFunction;
        public IMailAccount MirrorSource { get { return Account; } }
        public readonly ImapAccount Account;

        public ImapAccountControl(ImapAccount account)
        {
            InitializeComponent();
            Account = account;
            _dataContext = new ImapAccountDataContext(Account, RemoveImapMailboxFunction);
            DataContext = _dataContext;
        }

        private void UseGmailCheckBox_Checked(object sender, RoutedEventArgs e)
        {
            if (sender.Equals(UseGmailCheckBox) && UseGmailCheckBox.IsChecked == true)
            {
                _dataContext.SSL = true;
                _dataContext.Server = "imap.gmail.com";
            }
        }

        private void AddMailboxClick(object sender, RoutedEventArgs e)
        {
            Account.Mailboxes.Add<AuthenticatedMailbox>(new AuthenticatedMailbox(Account));
            Logger.Debug("Added a mailbox to account " + Account);
            //_dataContext.ImapMailboxes.Add(new ImapMailbox(this) { RemoveImapMailboxFunction = RemoveImapMailboxFunction});
        }

        private void RemoveImapMailboxFunction(object o)
        {
            var item = o as AuthenticatedMailbox;
            if (item != null)
            {
                if (Account.Mailboxes.Contains<AuthenticatedMailbox>(item))
                {
                    Account.Mailboxes.Remove<AuthenticatedMailbox>(item);
                }
            }
        }

        private void RemoveItem(object sender, RoutedEventArgs e)
        {
            if (RemoveAccountFunction != null)
            {
                RemoveAccountFunction(Account);
            }
        }

        public List<AuthenticatedMailbox> PendingMailboxes()
        {
            var pending = new List<AuthenticatedMailbox>();
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

    public class ImapAccountDataContext : INotifyPropertyChanged
    {
        private readonly ImapAccount _account;
        //public ObservableCollection<ImapMailbox> ImapMailboxes { get; set; }
        public AuthenticatedMailboxList Mailboxes { get; set; }
        public event PropertyChangedEventHandler PropertyChanged;

        public String Server
        {
            get { return _account.Server; }
            set { _account.Server = value; OnPropertyChanged("Server"); }
        }

        public bool SSL
        {
            get { return _account.UseSsl; }
            set { _account.UseSsl = value; OnPropertyChanged("SSL"); }
        }

        public ImapAccountDataContext(ImapAccount account, Action<AuthenticatedMailbox> removeMailboxAction)
        {
            _account = account;
            // ImapMailboxes = new ObservableCollection<ImapMailbox>();
            //ImapMailboxes.CollectionChanged += (sender, args) => OnPropertyChanged("ImapMailboxes");
            Mailboxes = new AuthenticatedMailboxList(removeMailboxAction, Dispatcher.CurrentDispatcher);
            Mailboxes.SetList(account.Mailboxes.UnderlyingCollection);
            Mailboxes.CollectionChanged += (sender, args) => OnPropertyChanged("Mailboxes");
        }

        [NotifyPropertyChangedInvocator]
        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChangedEventHandler handler = PropertyChanged;
            if (handler != null) handler(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
