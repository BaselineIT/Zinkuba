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
        private readonly ExchangeAccount _account;
        public IMailAccount Account { get { return _account; } }

        public ExchangeAccountControl(ExchangeAccount account)
        {
            InitializeComponent();
            _account = account;
            _dataContext = new ExchangeAccountDataContext(_account, RemoveMailboxFunction);
            DataContext = _dataContext;

        }

        private void RemoveItem(object sender, RoutedEventArgs e)
        {
            if (RemoveAccountFunction != null)
            {
                RemoveAccountFunction(_account);
            }
        }

        private void AddMailboxClick(object sender, RoutedEventArgs e)
        {
            _account.Mailboxes.Add<AuthenticatedMailbox>(new AuthenticatedMailbox(_account));
            Logger.Debug("Added a mailbox to account " + _account);
        }

        private void RemoveMailboxFunction(IMailbox o)
        {
            var item = o as AuthenticatedMailbox;
            if (item != null)
            {
                if (_account.Mailboxes.Contains<AuthenticatedMailbox>(item))
                {
                    _account.Mailboxes.Remove<AuthenticatedMailbox>(item);
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

        public IMailAccount MirrorSource { get { return _account; } }
    }


    public class ExchangeAccountDataContext : INotifyPropertyChanged
    {
        private readonly ExchangeAccount _account;
        public AuthenticatedMailboxList Mailboxes { get; set; }


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
