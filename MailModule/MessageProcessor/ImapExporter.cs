using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Sockets;
using System.Text.RegularExpressions;
using System.Threading;
using log4net;
using S22.Imap;
using Zinkuba.MailModule.API;
using Zinkuba.MailModule.MessageDescriptor;
using Exception = System.Exception;

namespace Zinkuba.MailModule.MessageProcessor
{
    public class ImapExporter : BaseMessageProcessor, IMessageWriter<RawMessageDescriptor>, IMessageSource
    {
        public IMessageReader<RawMessageDescriptor> NextReader { get; set; }
        private static readonly ILog Logger = LogManager.GetLogger(typeof (ImapExporter));
        private ImapClient _imapClient;
        private readonly string _server;
        private readonly bool _useSsl;
        private readonly string _username;
        private readonly string _password;
        public MailProvider Provider = MailProvider.DefaultImap;
        private Thread _imapThread;

        public int TotalMessages
        {
            get { return _totalMessages; }
            private set { _totalMessages = value; OnTotalMesssagesChanged(); }
        }

        private List<ImapMailbox> _mailBoxes;
        private int _totalMessages;
        private bool _testOnly;

        public ImapExporter(String username, String password, String server, bool useSSL)
        {
            _password = password;
            _server = server;
            _username = username;
            _useSsl = useSSL;
            Name = username;
        }

        public Type OutMessageDescriptorType()
        {
            return typeof(RawMessageDescriptor);
        }

        public override void Initialise()
        {
            Status = MessageProcessorStatus.Initialising;
            NextReader.Initialise();
            _mailBoxes = new List<ImapMailbox>();
            try
            {
                _imapClient = new ImapClient(_server, _useSsl ? 993 : 143, _username, _password, AuthMethod.Login, _useSsl,
                    delegate { return true; });
                Logger.Debug("Logged into " + _server + " as " + _username);
                var folders = _imapClient.ListMailboxes();
                foreach (var folderPath in folders)
                {
                    var destinationFolder = FolderMapping.ApplyMappings(folderPath, Provider);
                    if (!String.IsNullOrWhiteSpace(destinationFolder))
                    {
                        try
                        {
                            var folder = _imapClient.GetMailboxInfo(folderPath);
                            if (folder.Messages == 0)
                            {
                                Logger.Debug("Skipping folder " + folderPath + ", no messages.");
                                continue;
                            }
                            _mailBoxes.Add(new ImapMailbox() {MappedDestination = destinationFolder, Mailbox = folder});
                            TotalMessages += !_testOnly ? folder.Messages : (folder.Messages > 20 ? 20 : folder.Messages);
                        }
                        catch (Exception ex)
                        {
                            Logger.Error("Failed to get Mailbox " + folderPath + ", skipping.", ex);
                        }
                    }
                    else
                    {
                        Logger.Info("Ignoring folder " + folderPath + ", no destination specified.");
                    }
                }
            }
            catch (InvalidCredentialsException ex)
            {
                Logger.Error("Imap Runner for " + _username + " [********] to " + _server + " failed : " + ex.Message,ex);
                throw new MessageProcessorException("Imap Runner for " + _username + " [********] to " + _server + " failed : " + ex.Message) { Status = MessageProcessorStatus.AuthFailure };
            }
            catch (SocketException ex)
            {
                Logger.Error("Imap Runner for " + _username + " [********] to " + _server + " failed : " + ex.Message, ex);
                throw new MessageProcessorException("Imap Runner for " + _username + " [********] to " + _server + " failed : " + ex.Message) 
                { Status = MessageProcessorStatus.ConnectionError };
            }
            Status = MessageProcessorStatus.Initialised;
            Logger.Info("ExchangeExporter Initialised");
        }

        public void Start()
        {
            _imapThread = new Thread(RunImap) { IsBackground = true, Name = "imapThread-" + _username };
            _imapThread.Start();
        }

        public bool TestOnly
        {
            get { return _testOnly; }
            set { _testOnly = value; }
        }

        public event EventHandler TotalMessagesChanged;

        protected virtual void OnTotalMesssagesChanged()
        {
            EventHandler handler = TotalMessagesChanged;
            if (handler != null) handler(this, EventArgs.Empty);
        }

        private void RunImap()
        {
            try
            {
                Status = MessageProcessorStatus.Started;
                foreach (var mailbox in _mailBoxes)
                {
                    if (Closed)
                    {
                        Logger.Warn("Cancellation requested");
                        break;
                    }
                    var folder = mailbox.Mailbox;
                    var folderPath = folder.Name;
                    var destinationFolder = mailbox.MappedDestination;
                    try
                    {
                        var uids = _imapClient.Search(SearchCondition.All(), folderPath);
                        Logger.Info("Processing folder Name=" + folderPath + " Destination=" + destinationFolder +
                                    ", Count=" + folder.Messages + ", Unread=" + folder.Unread);
                        var count = folder.Messages;
                        var countQueued = 0;
                        foreach (var uid in uids.Reverse())
                        {
                            if (Closed)
                            {
                                Logger.Warn("Cancellation requested");
                                break;
                            }
                            if (_testOnly && count >= 20)
                            {
                                Logger.Debug("Testing only, hit 20 limit");
                                break;
                            }
                            try
                            {
                                var flags = _imapClient.GetMessageFlags(uid, folderPath).ToList();
                                if (flags.Contains(MessageFlag.Deleted))
                                {
                                    Logger.Debug("Skipping message " + folderPath + "/" + uid + ", its marked for deletion");
                                    IgnoredMessageCount++;
                                    continue;
                                }
                                var message = GetImapMessage(_imapClient, folder, uid);
                                message.Subject = Regex.Match(message.RawMessage, @"[\r\n]Subject: (.*?)[\r\n]").Groups[1].Value;
                                //Logger.Debug(folder + "/" + uid + "/" + subject + " " + String.Join(", ",flags));
                                if (!flags.Contains(MessageFlag.Seen))
                                {
                                    message.Flags.Add(MessageFlags.Unread);
                                }
                                if (flags.Contains(MessageFlag.Flagged))
                                {
                                    message.Flags.Add(MessageFlags.FollowUp);
                                }
                                message.SourceFolder = folderPath;
                                message.DestinationFolder = destinationFolder;
                                NextReader.Process(message);
                                SucceededMessageCount++;
                                countQueued++;
                            }
                            catch (Exception ex)
                            {
                                Logger.Error("Failed to get and enqueue Imap message [" + uid + "]", ex);
                                FailedMessageCount++;
                            }
                            ProcessedMessageCount++;
                        }
                        Logger.Info("Processed folder " + folderPath + ", read=" + count + ", queued=" + countQueued);
                    }
                    catch (Exception ex)
                    {
                        Logger.Error("Failed to select " + folderPath, ex);
                    }
                }
            }
            catch (Exception e)
            {
                Logger.Error("Failed to run exporter");
            }
            finally
            {
                Close();
            }
        }

        private RawMessageDescriptor GetImapMessage(ImapClient imapClient, MailboxInfo folder, uint uid)
        {
            String eml = imapClient.GetMessageData(uid, false, folder.Name);
            if (String.IsNullOrWhiteSpace(eml))
            {
                throw new IOException("Failed to read UID " + uid + " from imap server " + folder.Name);
            }
            return new RawMessageDescriptor { SourceId = uid.ToString(), RawMessage = eml };
        }

        public override void Close()
        {
            if (!Closed)
            {
                Closed = true;
                NextReader.Close();
                if (_imapClient != null)
                {
                    try
                    {
                        _imapClient.Logout();
                    }
                    catch (Exception ex)
                    {
                        Logger.Error("Failed to logout : " + ex.Message, ex);
                    }
                }
            }
        }
    }
}
