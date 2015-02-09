using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Reflection;
using System.Windows;
using log4net;
using Zinkuba.MailModule;
using Zinkuba.MailModule.API;
using Zinkuba.MailModule.MessageProcessor;

namespace OutlookTester
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private static readonly ILog Logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        public readonly ObservableCollection<UserAccount> UserAccounts;
        private PstTarget _pstWriter;
        private RawToMsgProcessor _rawToMsgProcessor;
        private ImapExporter _imapExporter;

        public MainWindow()
        {
            InitializeComponent();
            log4net.Config.XmlConfigurator.Configure(Assembly.GetExecutingAssembly().GetManifestResourceStream("Zinkuba.App.log4net.xml"));
#if DEBUG
            foreach (var manifestResourceName in Assembly.GetExecutingAssembly().GetManifestResourceNames())
            {
                Logger.Debug("Embedded Resource : " + manifestResourceName);
            }
#endif
            SaveFolder.Text = Environment.ExpandEnvironmentVariables(@"%USERPROFILE%\Downloads");
            UserAccounts = new ObservableCollection<UserAccount>();
            UserAccountsStackPanel.DataContext = UserAccounts;
            UserAccounts.Add(new UserAccount("test@baseline.cloud", "allmytest001"));
        }

        private void StartExport(object sender, RoutedEventArgs e)
        {
            foreach (var userAccount in UserAccounts)
            {
                if (!String.IsNullOrWhiteSpace(userAccount.Username) && !String.IsNullOrWhiteSpace(userAccount.Password))
                {
                    _pstWriter = new PstTarget(userAccount.Username + "@" + Server.Text, SaveFolder.Text);
                    _rawToMsgProcessor = new RawToMsgProcessor()
                    {
                        Name = userAccount.Username,
                        NextReader = _pstWriter
                    };
                    _imapExporter = new ImapExporter(userAccount.Username, userAccount.Password, Server.Text, SSL.IsChecked == true, TestOnlyCheckBox.IsChecked == true)
                    {
                        Provider = UseGmailCheckBox.IsChecked == true ? MailProvider.GmailImap : MailProvider.DefaultImap,
                        NextReader = _rawToMsgProcessor
                    };
                    var importer = new MessagePipeline(new List<IMessageProcessor>() { _imapExporter, _rawToMsgProcessor, _pstWriter });
                    userAccount.StartExporter(importer);
                }
            }
        }

        private void AddNewAccount(object sender, RoutedEventArgs e)
        {
            UserAccounts.Add(new UserAccount());
        }

        private void UseGmailCheckBox_Checked(object sender, RoutedEventArgs e)
        {
            if (sender.Equals(UseGmailCheckBox) && UseGmailCheckBox.IsChecked == true)
            {
                SSL.IsChecked = true;
                Server.Text = "imap.gmail.com";
            }
        }




    }
}
