using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Windows.Controls;
using log4net;
using Zinkuba.MailModule;
using Zinkuba.MailModule.API;
using Zinkuba.MailModule.MessageProcessor;

namespace Zinkuba.App
{
    /// <summary>
    /// Interaction logic for ImapMailbox.xaml
    /// </summary>
    public partial class ImapMailbox : UserControl, INotifyPropertyChanged
    {
        private readonly ImapAccountControl _account;
        private static readonly ILog Logger = LogManager.GetLogger(typeof(ImapMailbox));
        public string Username { get { return UsernameField.Text; } set { UsernameField.Text = value; } }
        public string Password { get { return PasswordField.Password; } set { PasswordField.Password = value; } }
        private MessagePipeline _exporter;
        private int _progress;
        public Action<object> RemoveImapMailboxFunction;



        public int IgnoredMails { get; set; }
        public int FailedMails { get; set; }
        public int ExportedMails { get; set; }

        public int Progress
        {
            get { return _progress; }
            set
            {
                if (_progress != value)
                {
                    _progress = value;
                    OnPropertyChanged("Progress");
                    OnPropertyChanged("ProgressText");
                    OnPropertyChanged("ExportedMails");
                    OnPropertyChanged("FailedMails");
                    OnPropertyChanged("IgnoredMails");
                }
            }
        }

        public String ProgressText
        {
            get { return _exporter == null ? "" : (_exporter.State == MessageProcessorStatus.Started ? "" + _progress + "%" : _exporter.State.ToString()); }
        }

        public MessagePipeline Exporter { get { return _exporter; } set { _exporter = value; } }

        public ImapMailbox(ImapAccountControl account)
        {
            InitializeComponent();
            _account = account;
            this.DataContext = this;
        }

        public void StartExporter(MessagePipeline exporter)
        {
            if (_exporter == null || _exporter.Failed)
            {
                _exporter = exporter;
                _exporter.SucceededMail += ExporterOnSucceededMail;
                _exporter.StateChanged += (sender, args) => OnPropertyChanged("ProgressText");
                try
                {
                    exporter.Start();
                }
                catch (Exception ex)
                {
                    Logger.Error("Failed to start exporter",ex);
                }
            }
        }

        private void ExporterOnSucceededMail(object sender, EventArgs eventArgs)
        {
            ExportedMails = _exporter.SucceededMails;
            FailedMails = _exporter.FailedMails;
            IgnoredMails = _exporter.IgnoredMails;
            Progress = (ExportedMails+FailedMails+IgnoredMails) * 100 / _exporter.TotalMails;
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChangedEventHandler handler = PropertyChanged;
            if (handler != null) handler(this, new PropertyChangedEventArgs(propertyName));
        }

        private void RemoveItem(object sender, System.Windows.RoutedEventArgs e)
        {
            if (RemoveImapMailboxFunction != null)
            {
                RemoveImapMailboxFunction(this);
            }
        }

        public IMessageSource GetSource()
        {
            return new ImapSource(UsernameField.Text, PasswordField.Password, _account.Server.Text, _account.Account.StartDate, _account.Account.EndDate,
                _account.SSL.IsChecked == true,_account.Account.LimitFolder == null ? null : new List<string> { _account.Account.LimitFolder })
            {
                Provider = _account.UseGmailCheckBox.IsChecked == true ? MailProvider.GmailImap : MailProvider.DefaultImap,
                Name = UsernameField.Text
            };
        }

        public bool Validate()
        {
            if (Dispatcher.CheckAccess())
            {
                return !String.IsNullOrWhiteSpace(UsernameField.Text) &&
                       !String.IsNullOrWhiteSpace(PasswordField.Password);
            }
            else
            {
                bool result = false;
                Dispatcher.Invoke(new Action(() => result = Validate()));
                return result;
            }
        }
    }
}
