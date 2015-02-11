using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Amib.Threading;
using log4net;
using Microsoft.Exchange.WebServices.Data;
using Zinkuba.MailModule.API;
using Zinkuba.MailModule.MessageDescriptor;
using Exception = System.Exception;

namespace Zinkuba.MailModule.MessageProcessor
{
    public class ExchangeTarget : BaseMessageProcessor, IMessageReader<RawMessageDescriptor>, IMessageDestination
    {
        private static readonly ILog Logger = LogManager.GetLogger(typeof(ExchangeTarget));

        private readonly string _hostname;
        private readonly string _username;
        private readonly string _password;
        private PCQueue<RawMessageDescriptor, ImportState> _queue;
        private ImportState _lastState;
        private ExchangeService service;
        private List<ExchangeFolder> folders;
        private readonly object _folderCreationLock = new object();
        private readonly SmartThreadPool pool;
        private ExtendedPropertyDefinition PR_MESSAGE_FLAGS_msgflag_read = new ExtendedPropertyDefinition(3591, MapiPropertyType.Integer);


        public ExchangeTarget(string hostname, string username, string password)
        {
            _hostname = hostname;
            _username = username;
            _password = password;
            pool = new SmartThreadPool() {MaxThreads = 5};
        }

        public void Process(RawMessageDescriptor message)
        {
            if (Status == MessageProcessorStatus.Started || Status == MessageProcessorStatus.Initialised)
            {
                Status = MessageProcessorStatus.Started;
                _queue.Consume(message);
            }
            else
            {
                FailedMessageCount++;
                throw new Exception("Cannot process, exchange target is not started.");
            }

        }

        public override void Initialise()
        {
            Status = MessageProcessorStatus.Initialising;
            service = ExchangeHelper.ExchangeConnect(_hostname, _username, _password);
            service.TraceEnabled = true;
            service.TraceFlags = TraceFlags.All;
            folders = new List<ExchangeFolder>();
            ExchangeHelper.GetAllFolders(service, new ExchangeFolder { Folder = Folder.Bind(service, WellKnownFolderName.MsgFolderRoot) }, folders, false);
            _lastState = new ImportState();
            _queue = new PCQueue<RawMessageDescriptor, ImportState>(Name + "-exchangeTarget")
            {
                ProduceMethod = ProcessMessage,
                InitialiseProducer = () => _lastState,
                ShutdownProducer = ShutdownQueue
            };
            _queue.Start();
            Status = MessageProcessorStatus.Initialised;
        }

        private void ShutdownQueue(ImportState state, Exception ex)
        {
            try
            {
                if (ex != null) throw ex;
                Status = MessageProcessorStatus.Completed;
            }
            catch (Exception e)
            {
                Logger.Error("Message Reader failed : " + e.Message, e);
                Status = MessageProcessorStatus.UnknownError;
            }
            Close();
        }

        private ImportState ProcessMessage(RawMessageDescriptor msg, ImportState importState)
        {
            pool.QueueWorkItem(() =>
            {
                var start = Environment.TickCount;
                try
                {
                    try
                    {
                        ImportIntoEWS(msg);
                        Logger.Debug("Imported message " + msg + " into " + _username + "@" + _hostname + " [" +
                                     (Environment.TickCount - start) + "ms]");
                    }
                    catch (ServiceLocalException e)
                    {
                        if (e.Message.Equals("The type of the object in the store (MeetingRequest) does not match that of the local object (Message)."))
                        {
                            // this is an error we can ignore, it just means we sent it as an email message, but its actually a meeting request, exchange imports it anyway
                            Logger.Debug("Imported meeting request " + msg + " into " + _username + "@" + _hostname + " [" +
                                         (Environment.TickCount - start) + "ms]");
                        }
                        else
                        {
                            throw e;
                        }
                    }
                    SucceededMessageCount++;
                }
                catch (Exception e)
                {
                    Logger.Error("Failed to import message " + msg + " into " + _username + "@" + _hostname + " [" + (Environment.TickCount-start) + "ms]" + " : " + e.Message,e);
                    FailedMessageCount++;
                }
                finally
                {
                    ProcessedMessageCount++;
                }
            });
            return importState;
        }

        private void ImportIntoEWS(RawMessageDescriptor msg)
        {
            EmailMessage item = new EmailMessage(service)
            {
                MimeContent = new MimeContent("UTF-8", Encoding.UTF8.GetBytes(msg.RawMessage))
            };
            String folder = msg.DestinationFolder;

            //Logger.Debug("Importing " + folder + @"\" + item.Subject + " (" + String.Join(", ", msg.Flags) + ") [" + String.Join(", ", msg.Categories) + "]");

            // This is required to be set one way or another, otherwise the message is marked as new (not delivered)
            item.IsRead = !msg.Flags.Contains(MessageFlags.Unread);
            // Set the defaults to no receipts, to be corrected later by flags.
            item.IsReadReceiptRequested = false;
            item.IsDeliveryReceiptRequested = false;

            foreach (var messageFlag in msg.Flags)
            {
                try
                {
                    switch (messageFlag)
                    {
                        case MessageFlags.Associated:
                            {
                                item.IsAssociated = true;
                                break;
                            }
                        case MessageFlags.FollowUp:
                            {
                                item.Flag = new Flag() { FlagStatus = ItemFlagStatus.Flagged };
                                break;
                            }
                        case MessageFlags.ReminderSet:
                            {
                                item.IsReminderSet = true;
                                break;
                            }
                        case MessageFlags.ReadReceiptRequested:
                            {
                                // If the item is read already, we don't set a read receipt request as it will send one on save.
                                if (!item.IsRead)
                                {
                                    item.IsReadReceiptRequested = true;
                                }
                                break;
                            }
                        case MessageFlags.DeliveryReceiptRequested:
                            {
                                // this causes spam
                                //item.IsDeliveryReceiptRequested = true;
                                break;
                            }
                    }
                }
                catch (Exception e)
                {
                    Logger.Warn("Failed to set flag on " + folder + @"\" + item.Subject + ", ignoring flag.", e);
                }
            }
            try
            {
                if (!msg.Flags.Contains(MessageFlags.Draft))
                {
                    item.SetExtendedProperty(PR_MESSAGE_FLAGS_msgflag_read, 1);
                }
                if (msg.Importance != null) item.Importance = (Importance)msg.Importance;
                if (msg.Sensitivity != null) item.Sensitivity = (Sensitivity)msg.Sensitivity;
                if (msg.ReminderDueBy != null) item.ReminderDueBy = (DateTime)msg.ReminderDueBy;
                item.Categories = new StringList(msg.Categories);
            }
            catch (Exception e)
            {
                Logger.Warn(
                    "Failed to set metadata on " + folder + @"\" + item.Subject + ", ignoring metadata.", e);
            }

            FolderId parentFolderId = GetCreateFolder(msg.DestinationFolder);
            item.Save(parentFolderId);
        }

        private FolderId GetCreateFolder(string destinationFolder)
        {
            lock (_folderCreationLock)
            {
                if (!folders.Any(folder => folder.FolderPath.Equals(destinationFolder)))
                {
                    // folder doesn't exist
                    // getCreate its parent
                    var parentPath = Regex.Replace(destinationFolder, @"\\[^\\]+$", "");
                    FolderId parentFolderId = null;
                    if (parentPath.Equals(destinationFolder))
                    {
                        // we are at the root
                        parentPath = "";
                        parentFolderId = WellKnownFolderName.MsgFolderRoot;
                    }
                    else
                    {
                        parentFolderId = GetCreateFolder(parentPath);
                    }
                    Logger.Debug("Folder " + destinationFolder + " doesn't exist, creating.");
                    var destinationFolderName = Regex.Replace(destinationFolder, @"^.*\\", "");
                    Folder folder = new Folder(service) {DisplayName = destinationFolderName};
                    folder.Save(parentFolderId);
                    folders.Add(new ExchangeFolder()
                    {
                        Folder = folder,
                        FolderId = folder.Id,
                        FolderPath =
                            String.IsNullOrEmpty(parentPath)
                                ? destinationFolderName
                                : parentPath + @"\" + destinationFolderName,
                    });
                    return folder.Id;
                }
                else
                {
                    return folders.First(folder => folder.FolderPath.Equals(destinationFolder)).FolderId;
                }
            }
        }

        public Type InMessageDescriptorType()
        {
            return typeof(RawMessageDescriptor);
        }

        public override void Close()
        {
            if (!Closed && Status != MessageProcessorStatus.Idle)
            {
                Closed = true;
                if (_queue != null) { _queue.Close(); }
                try
                {
                    // nothing to do to close ews, its rest
                }
                catch
                {
                    Logger.Error("Failed to close ews");
                }
                foreach (var messageStats in _lastState.Stats)
                {
                    Logger.Info(messageStats.Value.ToString());
                }
            }
        }

        public List<string> ImportedIds { get; private set; }
    }
}
