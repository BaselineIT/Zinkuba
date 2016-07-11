using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using log4net;
using Zinkuba.App.Annotations;
using Zinkuba.App.MailAccount;
using Zinkuba.MailModule;
using Zinkuba.MailModule.API;
using Zinkuba.MailModule.MessageProcessor;

namespace Zinkuba.App.Mailbox
{
    public class AuthenticatedMailbox : IMailbox, INotifyPropertyChanged
    {
        private static readonly ILog Logger = LogManager.GetLogger(typeof (AuthenticatedMailbox));
        private readonly ServerMailAccount _account;
        public string Password;
        public MessagePipeline Exporter { get; private set; }
        public string Username;
        private string _id;
        public event EventHandler ExportStarted;

        protected virtual void OnExportStarted()
        {
            EventHandler handler = ExportStarted;
            if (handler != null) handler(this, EventArgs.Empty);
        }

        public AuthenticatedMailbox(ServerMailAccount account)
        {
            _account = account;
        }

        public IMessageSource GetSource()
        {
            if (_account is ExchangeAccount)
                return new ExchangeExporter(Username, Password, _account.Server, _account.StartDate,
                    _account.EndDate.AddDays(1),
                    String.IsNullOrWhiteSpace(_account.LimitFolder) ? null : new List<string>() {_account.LimitFolder})
                {
                    IncludePublicFolders = ((ExchangeAccount) _account).IncludePublicFolders
                };
            if (_account is ImapAccount)
                return new ImapExporter(Username, Password, _account.Server, _account.StartDate, _account.EndDate.AddDays(1), ((ImapAccount)_account).UseSsl, String.IsNullOrWhiteSpace(_account.LimitFolder) ? null : new List<string>() { _account.LimitFolder })
                {
                    Provider = ((ImapAccount)_account).Gmail ? MailProvider.GmailImap : MailProvider.DefaultImap,
                    PublicFolderRoot = ((ImapAccount)_account).IncludePublicFolders ? ((ImapAccount)_account).PublicFolderRoot : "",
                    IncludePublicFolders = ((ImapAccount)_account).IncludePublicFolders,
                };
            throw new Exception("Unknown Account");
        }

        public void StartExporter(MessagePipeline exporter)
        {
            if (Exporter == null || Exporter.Failed)
            {
                Exporter = exporter;
                OnExportStarted();
                try
                {
                    exporter.Start();
                }
                catch (Exception ex)
                {
                    Logger.Error("Failed to start exporter", ex);
                }
            }
        }

        public string Id
        {
            get { return _id; }
            set
            {
                if (value == _id) return;
                _id = value;
                OnPropertyChanged();
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        [NotifyPropertyChangedInvocator]
        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChangedEventHandler handler = PropertyChanged;
            if (handler != null) handler(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
