using System;
using System.ComponentModel;
using System.Windows.Controls;
using log4net;
using Rendezz.UI;
using Zinkuba.App.MailAccount;
using Zinkuba.App.Mailbox;
using Zinkuba.MailModule;
using Zinkuba.MailModule.API;

namespace Zinkuba.App
{
    /// <summary>
    /// Interaction logic for AuthenticatedMailboxControl.xaml
    /// </summary>
    public partial class AuthenticatedMailboxControl : UserControl, IReflectedObject<AuthenticatedMailbox>
    {
        public readonly AuthenticatedMailbox Mailbox;
        private static readonly ILog Logger = LogManager.GetLogger(typeof(AuthenticatedMailboxControl));
        public Action<AuthenticatedMailbox> RemoveMailboxFunction;
        private readonly AuthenticatedMailboxData _authenticatedMailboxData;

        public AuthenticatedMailboxControl(AuthenticatedMailbox mailbox)
        {
            InitializeComponent();
            Mailbox = mailbox;
            Mailbox.ExportStarted += MailboxOnExportStarted;
            _authenticatedMailboxData = new AuthenticatedMailboxData(Mailbox);
            DataContext = _authenticatedMailboxData;
        }

        private void MailboxOnExportStarted(object sender, EventArgs eventArgs)
        {
            _authenticatedMailboxData.Exporter = Mailbox.Exporter;
            Mailbox.Exporter.ExportedMail += ExporterOnExportedMail;
            Mailbox.Exporter.StateChanged += (sender1, args) => _authenticatedMailboxData.OnPropertyChanged("ProgressText");
        }

        private void RemoveItem(object sender, System.Windows.RoutedEventArgs e)
        {
            if (RemoveMailboxFunction != null)
            {
                RemoveMailboxFunction(Mailbox);
            }
        }

        private void ExporterOnExportedMail(object sender, EventArgs eventArgs)
        {
            _authenticatedMailboxData.ExportedMails = Mailbox.Exporter.ExportedMails;
            _authenticatedMailboxData.FailedMails = Mailbox.Exporter.FailedMails;
            _authenticatedMailboxData.IgnoredMails = Mailbox.Exporter.IgnoredMails;
            _authenticatedMailboxData.Progress = (_authenticatedMailboxData.ExportedMails + _authenticatedMailboxData.FailedMails + _authenticatedMailboxData.IgnoredMails) * 100 / Mailbox.Exporter.TotalMails;
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

        public AuthenticatedMailbox MirrorSource { get { return Mailbox; } }

        private void PasswordField_LostFocus(object sender, System.Windows.RoutedEventArgs e)
        {
            _authenticatedMailboxData.Password = PasswordField.Password;
        }
    }

    public class AuthenticatedMailboxData : MailboxDataContext
    {

        private readonly AuthenticatedMailbox _authenticatedMailbox;

        public String Username
        {
            get { return _authenticatedMailbox.Username; }
            set { _authenticatedMailbox.Username = value; }
        }
        public String Password
        {
            get { return _authenticatedMailbox.Password; }
            set { _authenticatedMailbox.Password = value; }
        }

        public String MailboxId { get { return _authenticatedMailbox.Id; } set { _authenticatedMailbox.Id = value; OnPropertyChanged("MailboxId"); } }

        public AuthenticatedMailboxData(AuthenticatedMailbox authenticatedMailbox)
        {
            _authenticatedMailbox = authenticatedMailbox;
            _authenticatedMailbox.PropertyChanged +=
                (sender, args) => { if (args.PropertyName.Equals("Id")) OnPropertyChanged("MailboxId"); };
        }



    }
}
