using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Threading;
using log4net;
using log4net.Repository.Hierarchy;
using Rendezz.UI;
using Zinkuba.App;
using Zinkuba.App.Annotations;
using Zinkuba.App.Folder;
using Zinkuba.App.Mailbox;
using Zinkuba.MailModule.API;
using UserControl = System.Windows.Controls.UserControl;

namespace Zinkuba.App
{
    /// <summary>
    /// Interaction logic for MboxAccount.xaml
    /// </summary>
    public partial class MboxMailboxControl : UserControl, IReflectedObject<MboxMailbox>
    {
        private static readonly ILog Logger = LogManager.GetLogger(typeof (MboxMailboxControl));
        private MboxMailboxDataContext _dataContext;

        public MboxMailbox Mailbox
        {
            get { return _mailbox; }
        }

        public Action<MboxMailbox> RemoveMailboxFunction;
        private MboxMailbox _mailbox;

        public MboxMailboxControl(MboxMailbox sourceMailbox)
        {
            InitializeComponent();
            _mailbox = sourceMailbox;
            _dataContext = new MboxMailboxDataContext(_mailbox,RemoveFolderFunction);
            Mailbox.ExportStarted += MailboxOnExportStarted;
            DataContext = _dataContext;
        }

        private void MailboxOnExportStarted(object sender, EventArgs eventArgs)
        {
            _dataContext.Exporter = Mailbox.Exporter;
            Mailbox.Exporter.ExportedMail += ExporterOnExportedMail;
            Mailbox.Exporter.FailedMail += ExporterOnFailedMail;
            Mailbox.Exporter.IgnoredMail += ExporterOnIgnoredMail;
            Mailbox.Exporter.StateChanged += (sender1, args) => _dataContext.OnPropertyChanged("ProgressText");
        }

        private void ExporterOnIgnoredMail(object sender, EventArgs eventArgs)
        {
            _dataContext.IgnoredMails = Mailbox.Exporter.IgnoredMails;
            _dataContext.Progress = 0;
        }

        private void ExporterOnFailedMail(object sender, EventArgs eventArgs)
        {
            _dataContext.FailedMails = Mailbox.Exporter.FailedMails;
            _dataContext.Progress = 0;
        }

        private void ExporterOnExportedMail(object sender, EventArgs eventArgs)
        {
            _dataContext.ExportedMails = Mailbox.Exporter.ExportedMails;
            _dataContext.Progress = 0;
        }
        private void AddMboxClick(object sender, RoutedEventArgs e)
        {
            Mailbox.Folders.Add(new MboxFolder());
            Logger.Debug("Added new folder to " + Mailbox.Name);
        }

        private void RemoveFolderFunction(MboxFolder item)
        {
            if (item != null)
            {
                if (Mailbox.Folders.Contains(item))
                {
                    Mailbox.Folders.Remove(item);
                }
            }
        }

        private void RemoveItem(object sender, RoutedEventArgs e)
        {
            if (RemoveMailboxFunction != null)
            {
                RemoveMailboxFunction(_mailbox);
            }
        }

        public MboxMailbox MirrorSource { get { return Mailbox; } }

        public bool Validate()
        {
            return !String.IsNullOrWhiteSpace(Mailbox.Name) && Mailbox.Folders.All(mboxFolder => !String.IsNullOrWhiteSpace(mboxFolder.FolderPath) && !String.IsNullOrWhiteSpace(mboxFolder.MboxPath));
        }

        private void AddMboxFolderClick(object sender, RoutedEventArgs e)
        {
            var window = new OpenFolderDialog();
            window.Owner = MainWindow.CurrentInstance;
            window.ShowDialog();
            if (window.Folder != null && Directory.Exists(window.Folder))
            {
                foreach (var enumerateFile in Directory.EnumerateFiles(window.Folder,"*mbox",SearchOption.AllDirectories))
                {
                    Mailbox.Folders.Add(new MboxFolder() { MboxPath = enumerateFile });
                }
            }
        }
    }

    public class MboxMailboxDataContext : MailboxDataContext
    {
        private readonly MboxMailbox _mailbox;
        public MboxFolderList FolderList { get; set; }

        public String Name
        {
            get { return _mailbox.Name; }
            set { _mailbox.Name = value; }
        }

        public new String ProgressText
        {
            get { return Exporter == null ? "" : (Exporter.State == MessageProcessorStatus.Started ? "No Progress Info" : Exporter.State.ToString()); }
        }

        public MboxMailboxDataContext(MboxMailbox mailbox, Action<MboxFolder> removeFolderAction)
        {
            _mailbox = mailbox;
            FolderList = new MboxFolderList(removeFolderAction,Dispatcher.CurrentDispatcher);
            FolderList.SetList(mailbox.Folders);
        }

    }
}
