using System;
using System.ComponentModel;
using System.Windows.Controls;
using log4net;
using S22.Imap;
using Zinkuba.MailModule;
using Zinkuba.MailModule.API;
using Zinkuba.MailModule.MessageProcessor;

namespace OutlookTester
{
    /// <summary>
    /// Interaction logic for UserAccount.xaml
    /// </summary>
    public partial class UserAccount : UserControl, INotifyPropertyChanged
    {
        private static readonly ILog Logger = LogManager.GetLogger(typeof(UserAccount));
        public string Username { get { return UsernameField.Text; } set { UsernameField.Text = value; } }
        public string Password { get { return PasswordField.Password; } set { PasswordField.Password = value; } }
        private MessagePipeline _exporter;
        private int _progress;



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

        public UserAccount()
        {
            InitializeComponent();
            this.DataContext = this;
        }

        public UserAccount(string username, string password)
            : this()
        {
            Username = username;
            Password = password;
        }

        public void StartExporter(MessagePipeline exporter)
        {
            if (_exporter == null || _exporter.Failed)
            {
                _exporter = exporter;
                _exporter.ExportedMail += ExporterOnExportedMail;
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

        private void ExporterOnExportedMail(object sender, EventArgs eventArgs)
        {
            ExportedMails = _exporter.ExportedMails;
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
    }
}
