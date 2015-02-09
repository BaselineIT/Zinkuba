using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media.Animation;
using System.Windows.Threading;
using log4net;
using Rendezz.UI;
using Zinkuba.App.Annotations;
using Zinkuba.App.MailAccount;
using Zinkuba.App.Mailbox;
using Zinkuba.MailModule;
using Zinkuba.MailModule.API;
using Zinkuba.MailModule.MessageDescriptor;
using Zinkuba.MailModule.MessageProcessor;

namespace Zinkuba.App
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private static readonly ILog Logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        private readonly ZinkubaDataContext _dataContext;
        private Collection<MessagePipeline> _pipelines;
        private ObservableCollection<IMailAccount> MailSources;
        private static Collection<IMessageConnector> _connectorTypes;
        public object StartExporterLock = new object();

        private PstDestinationControl _pstDestinationControl;
        private ExchangeDestinationControl _exchangeDestinationControl;
        private int idIndex = 0;
        private readonly List<String> _removedMailboxes;

        internal static MainWindow CurrentInstance { get; set; }

        public MainWindow()
        {
            InitializeComponent();
            CurrentInstance = this;
            log4net.Config.XmlConfigurator.Configure(Assembly.GetExecutingAssembly().GetManifestResourceStream("Zinkuba.App.log4net.xml"));
#if DEBUG
            foreach (var manifestResourceName in Assembly.GetExecutingAssembly().GetManifestResourceNames())
            {
                Logger.Debug("Embedded Resource : " + manifestResourceName);
            }
#endif
            _pipelines = new Collection<MessagePipeline>();
            MailSources = new ObservableCollection<IMailAccount>();
            _dataContext = new ZinkubaDataContext(MailSources, RemoveMailSource);
            DataContext = _dataContext;
            _removedMailboxes = new List<string>();
            TargetComboBox.SelectedIndex = 0;
        }

        private void StartExport(object sender, RoutedEventArgs e)
        {
            RunExporters();
        }

        private void RunExporters()
        {
            lock (StartExporterLock)
            {
                foreach (var mailSource in _dataContext.MailSources)
                {
                    if (mailSource is ImapAccountControl)
                    {
                        Logger.Debug("Processing Imap account " + mailSource);
                        var mailAccountControl = (ImapAccountControl)mailSource;
                        foreach (var mailBox in mailAccountControl.PendingMailboxes())
                        {
                            StartExporter(mailBox);
                        }
                    }
                    else if (mailSource is ExchangeAccountControl)
                    {
                        Logger.Debug("Processing Exchange account " + mailSource);
                        var mailAccountControl = (ExchangeAccountControl)mailSource;
                        foreach (var mailBox in mailAccountControl.PendingMailboxes())
                        {
                            StartExporter(mailBox);
                        }
                    }
                    else if (mailSource is MboxAccountControl)
                    {
                        Logger.Debug("Processing Mbox account " + mailSource);
                        var mailAccountControl = (MboxAccountControl)mailSource;
                        foreach (var mailBox in mailAccountControl.PendingMailboxes())
                        {
                            StartExporter(mailBox);
                        }
                    }
                }
            }
        }

        private void StartExporter(IMailbox mailbox)
        {
            var destinationControl = MailDestination.Content as IDestinationManager;
            if (destinationControl != null)
            {
                var target = destinationControl.GetDestination(mailbox.Id);
                if (target is IMessageReader)
                {
                    var source = mailbox.GetSource();
                    source.TestOnly = TestOnlyCheckBox.IsChecked == true;
                    target.Name = source.Name;
                    ThreadPool.QueueUserWorkItem(state =>
                    {
                        try
                        {
                            var exporter = ConnectSourceDestination(source, target);
                            mailbox.StartExporter(exporter);
                        }
                        catch (Exception ex)
                        {
                            Logger.Error("Failed to setup exporter " + source.Name, ex);
                        }
                    });
                }
            }
        }


        private MessagePipeline ConnectSourceDestination(IMessageSource source, IMessageDestination target)
        {
            var sourceWriter = source as IMessageWriter;
            var targetReader = target as IMessageReader;
            var pipelineElements = new List<IMessageProcessor>();
            if (sourceWriter != null && targetReader != null)
            {
                pipelineElements.Insert(0, target);
                if (sourceWriter.OutMessageDescriptorType() == typeof(RawMessageDescriptor))
                {
                    if (targetReader.InMessageDescriptorType() == typeof(RawMessageDescriptor))
                    {
                        // nothing to do here, join will go ok
                    }
                    else if (targetReader.InMessageDescriptorType() == typeof(MsgDescriptor))
                    {
                        pipelineElements.Insert(0,
                            new RawToMsgProcessor()
                            {
                                Name = source.Name,
                                NextReader = (IMessageReader<MsgDescriptor>)target
                            });
                    }
                    else
                    {
                        throw new Exception("Unsupported connection types, need " +
                                            targetReader.InMessageDescriptorType() + " writer.");
                    }
                }
                else if (sourceWriter.OutMessageDescriptorType() == typeof(MsgDescriptor))
                {
                    if (targetReader.InMessageDescriptorType() == typeof(MsgDescriptor))
                    {
                        // nothing to do here, join will go ok
                    }
                    else
                    {
                        throw new Exception("Unsupported connection types, need " +
                                            targetReader.InMessageDescriptorType() + " writer.");
                    }
                }
                else
                {
                    throw new Exception("Unsupported connection types, need " + sourceWriter.OutMessageDescriptorType() +
                                        " reader.");
                }
                pipelineElements.Insert(0, source);
            }
            return new MessagePipeline(pipelineElements);
        }

        private void AddAccountClick(object sender, RoutedEventArgs e)
        {
            if (SourceComboBox.SelectedItem != null)
            {
                IMailAccount account = null;
                if ("Imap".Equals(((ComboBoxItem)(SourceComboBox.SelectedItem)).Content))
                {
                    account = new ImapAccount();
                }
                else if ("Mbox".Equals(((ComboBoxItem)(SourceComboBox.SelectedItem)).Content))
                {
                    account = new MboxAccount();
                }
                else if ("Exchange".Equals(((ComboBoxItem)(SourceComboBox.SelectedItem)).Content))
                {
                    account = new ExchangeAccount();
                }
                else
                {
                    Logger.Warn("Unsupported Account '" + ((ComboBoxItem)(SourceComboBox.SelectedItem)).Content + "'");
                }
                if (account != null)
                {
                    account.AddedMailbox += AddMailboxEvent;
                    account.RemovedMailbox += RemoveMailboxEvent;
                    MailSources.Add(account);
                }
            }
        }

        private void RemoveMailboxEvent(object sender, IMailbox mailbox)
        {
            String id = mailbox.Id;
            RemoveMailboxById(id);
        }

        private void RemoveMailboxById(string id)
        {
            var destinationControl = MailDestination.Content as IDestinationManager;
            if (destinationControl != null)
            {
                destinationControl.RemoveDestination(id);
                if (!_removedMailboxes.Contains(id))
                {
                    _removedMailboxes.Add(id);
                }
            }
        }

        private void AddMailboxEvent(object sender, IMailbox mailbox)
        {
            var destinationControl = MailDestination.Content as IDestinationManager;
            if (destinationControl != null)
            {
                if (mailbox.Id == null)
                {
                    mailbox.Id = GetNextId();
                }
                destinationControl.AddDestination(mailbox.Id);
            }
        }

        public String GetNextId()
        {
            return (++idIndex).ToString();
        }

        private void RemoveMailSource(IMailAccount item)
        {
            if (item != null)
            {
                if (MailSources.Contains(item))
                {
                    int mbCount = item.Mailboxes.Count;
                    if (mbCount > 0)
                    {

                        for (int i = mbCount - 1; i >= 0; i--)
                        {
                            RemoveMailboxById(item.Mailboxes[i].Id);
                        }
                    }
                    MailSources.Remove(item);
                    Logger.Debug("Removed mailbox " + item);
                }
            }
        }

        private void TargetComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            SetTarget();
        }

        private void SetTarget()
        {
            if (TargetComboBox.SelectedItem != null)
            {
                if ("Exchange".Equals(((ComboBoxItem)(TargetComboBox.SelectedItem)).Content))
                {
                    if (_exchangeDestinationControl == null)
                    {
                        _exchangeDestinationControl = new ExchangeDestinationControl();
                    }
                    MailDestination.Content = _exchangeDestinationControl;
                }
                else if ("PST".Equals(((ComboBoxItem)(TargetComboBox.SelectedItem)).Content))
                {
                    if (_pstDestinationControl == null)
                    {
                        _pstDestinationControl = new PstDestinationControl();
                    }
                    MailDestination.Content = _pstDestinationControl;
                }
                PopulateMailboxesInDestination();
            }
        }

        private void PopulateMailboxesInDestination()
        {
            foreach (var mailAccount in MailSources)
            {
                for (int i = 0; i < mailAccount.Mailboxes.Count; i++)
                {
                    AddMailboxEvent(this, mailAccount.Mailboxes[i]);
                }
            }
            foreach (var removedMailbox in _removedMailboxes)
            {
                RemoveMailboxById(removedMailbox);
            }
        }
    }

    public class ZinkubaDataContext : INotifyPropertyChanged
    {
        public MailSourceList MailSources { get; set; }
        //public ObservableCollection<UserControl> MailSources { get; set; }

        public ZinkubaDataContext(ObservableCollection<IMailAccount> sources, Action<IMailAccount> removeMailAccountFunction)
        {
            MailSources = new MailSourceList(removeMailAccountFunction, Dispatcher.CurrentDispatcher);
            MailSources.SetList(sources);
            MailSources.CollectionChanged += (sender, args) => OnPropertyChanged("MailSources");
        }

        public event PropertyChangedEventHandler PropertyChanged;

        [NotifyPropertyChangedInvocator]
        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChangedEventHandler handler = PropertyChanged;
            if (handler != null) handler(this, new PropertyChangedEventArgs(propertyName));
        }

    }

    public class MailSourceList : ObservableListMirror<IMailAccount, IMailAccountControl>
    {
        private readonly Action<IMailAccount> _removeMailSource;

        public MailSourceList(Action<IMailAccount> removeMailSource, Dispatcher dispatcher)
            : base(dispatcher)
        {
            _removeMailSource = removeMailSource;
        }


        protected override IMailAccountControl CreateNew(IMailAccount sourceObject)
        {
            if (sourceObject is ExchangeAccount)
            {
                return new ExchangeAccountControl((ExchangeAccount)sourceObject) { RemoveAccountFunction = _removeMailSource };
            }
            else if (sourceObject is ImapAccount)
            {
                return new ImapAccountControl((ImapAccount)sourceObject) { RemoveAccountFunction = _removeMailSource };
            }
            else if (sourceObject is MboxAccount)
            {
                return new MboxAccountControl((MboxAccount)sourceObject) { RemoveAccountFunction = _removeMailSource };
            }
            throw new Exception("Unknown account");
        }
    }
}
