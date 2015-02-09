using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;
using log4net;
using log4net.Repository.Hierarchy;
using Zinkuba.App.Annotations;
using Zinkuba.App.MailAccount;
using Zinkuba.App.Mailbox;

namespace Zinkuba.App
{
    /// <summary>
    /// Interaction logic for ExchangeAccountControl.xaml
    /// </summary>
    public partial class ExchangeAccountControl : UserControl, IMailAccountControl
    {
        private static readonly ILog Logger = LogManager.GetLogger(typeof (ExchangeAccountControl));
        public Action<IMailAccount> RemoveAccountFunction;
        private readonly ExchangeAccountDataContext _dataContext;
        public readonly ExchangeAccount Account;

        public ExchangeAccountControl(ExchangeAccount account)
        {
            InitializeComponent();
            Account = account;
            _dataContext = new ExchangeAccountDataContext(Account, RemoveMailboxFunction);
            DataContext = _dataContext;
            _dataContext.StartDate = DateTime.Now.AddYears(-10);
            _dataContext.EndDate = DateTime.Now.Date;
        }

        private void RemoveItem(object sender, RoutedEventArgs e)
        {
            if (RemoveAccountFunction != null)
            {
                RemoveAccountFunction(Account);
            }
        }

        private void AddMailboxClick(object sender, RoutedEventArgs e)
        {
            Account.Mailboxes.Add<AuthenticatedMailbox>(new AuthenticatedMailbox(Account));
            Logger.Debug("Added a mailbox to account " + Account);
        }

        private void RemoveMailboxFunction(IMailbox o)
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

        public IMailAccount MirrorSource { get { return Account; } }
    }


    public class ExchangeAccountDataContext : INotifyPropertyChanged
    {
        private readonly ExchangeAccount _account;
        private DateTime _startDate;
        private DateTime _endDate;
        public AuthenticatedMailboxList Mailboxes { get; set; }

        public DateTime StartDate
        {
            get { return _account.StartDate; }
            set
            {
                if (value.Equals(_account.StartDate)) return;
                _account.StartDate = value;
                OnPropertyChanged("StartDate");
            }
        }

        public DateTime EndDate
        {
            get { return _account.EndDate; }
            set
            {
                if (value.Equals(_account.EndDate)) return;
                _account.EndDate = value;
                OnPropertyChanged("EndDate");
            }
        }

        public String Server
        {
            get { return _account.Server; } 
            set { _account.Server = value; }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        public ExchangeAccountDataContext(ExchangeAccount account, Action<AuthenticatedMailbox> removeMailboxAction)
        {
            _account = account;
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
