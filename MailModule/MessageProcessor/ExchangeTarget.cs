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
using Microsoft.Office.Interop.Outlook;
using Zinkuba.MailModule.API;
using Zinkuba.MailModule.MessageDescriptor;
using Exception = System.Exception;
using Folder = Microsoft.Exchange.WebServices.Data.Folder;

namespace Zinkuba.MailModule.MessageProcessor
{
    public class ExchangeTarget : BaseMessageProcessor, IMessageReader<RawMessageDescriptor>, IMessageDestination
    {
        private const int MaxBufferSize = 5242880;
        private const int MaxBufferMails = 15;
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
        private String previousFolder;
        private List<Item> itemBuffer;
        private int bufferedSize;


        public ExchangeTarget(string hostname, string username, string password)
        {
            _hostname = hostname;
            _username = username;
            _password = password;
            itemBuffer = new List<Item>();
            pool = new SmartThreadPool() { MaxThreads = 5 };
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
            //service.TraceEnabled = true;
            //service.TraceFlags = TraceFlags.All;
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
            //pool.QueueWorkItem(() => ImportRunner(msg));
            ImportBulk(msg);
            return importState;
        }

        private void ImportBulk(RawMessageDescriptor msg)
        {
            if ((!String.IsNullOrEmpty(previousFolder) && !previousFolder.Equals(msg.DestinationFolder)) || bufferedSize > MaxBufferSize || itemBuffer.Count > MaxBufferMails)
            {
                CommitBufferedMessages();
            }
            try
            {
                Logger.Debug("Buffering " + msg.SourceId + " : " + msg.DestinationFolder + @"\" + msg.Subject + " (" + String.Join(", ", msg.Flags) + ") [" + msg.ItemClass + "]");
                EmailMessage item = PrepareEWSItem(msg);
                itemBuffer.Add(item);
                //SucceededMessageCount++;
                bufferedSize += msg.RawMessage.Length;
            }
            catch (Exception e)
            {
                Logger.Error("Failed to prepare item " + msg.SourceId + " : " + msg.DestinationFolder + @"\" + msg.Subject + " for import ", e);
                FailedMessageCount++;
                ProcessedMessageCount++;
            }
            //ProcessedMessageCount++;
            previousFolder = msg.DestinationFolder;
        }

        private void CommitBufferedMessages()
        {
            if (itemBuffer.Count > 0)
            {
                var start = Environment.TickCount;
                FolderId parentFolderId = GetCreateFolder(previousFolder);
                try
                {
                    var response = service.CreateItems(itemBuffer.AsEnumerable(), parentFolderId,
                        MessageDisposition.SaveOnly, null);
                    if (response.OverallResult != ServiceResult.Success)
                    {
                        var count = 0;
                        var successCount = 0;
                        foreach (var serviceResponse in response)
                        {
                            if (serviceResponse.Result == ServiceResult.Success)
                            {
                                successCount++;
                                SucceededMessageCount++;
                            }
                            else
                            {
                                Logger.Error("Failed to import message " + itemBuffer[count] + "[" + itemBuffer[count].ItemClass + "] into " + _username + "@" + _hostname + "/" +
                                             previousFolder + " [" + (Environment.TickCount - start) + "ms]" +
                                             " : [" + serviceResponse.ErrorCode + "]" + serviceResponse.ErrorMessage);
                                FailedMessageCount++;
                            }
                            count++;
                        }
                        Logger.Info("Saved " + successCount + " of a possible " + itemBuffer.Count + " messages into " +
                                    _username + "@" + _hostname + "/" + previousFolder + " [" +
                                    (Environment.TickCount - start) + "ms]");
                    }
                    else
                    {
                        Logger.Info("Saved " + itemBuffer.Count + " messages [" + (bufferedSize/1024) + "Kb] into " +
                                    _username + "@" + _hostname + "/" + previousFolder + " [" +
                                    (Environment.TickCount - start) + "ms]");
                        SucceededMessageCount += itemBuffer.Count;
                    }
                }
                // This exception still gets through, its very odd, this should be part of the response codes
                catch (ServiceLocalException e)
                {
                    if (Regex.Match(e.Message,@"The type of the object in the store \(\w+\) does not match that of the local object \(\w+\).").Success)
                    {
                        // this is an error we can ignore, it just means we sent it as an email message, but its actually a meeting request/response, exchange imports it anyway
                        Logger.Info("Saved " + itemBuffer.Count + " messages [" + (bufferedSize / 1024) + "Kb] into " +
                                    _username + "@" + _hostname + "/" + previousFolder + " [" +
                                    (Environment.TickCount - start) + "ms]");
                        SucceededMessageCount += itemBuffer.Count;
                    }
                    else
                    {
                        throw e;
                    }
                }
                catch (Exception e)
                {
                    Logger.Error("Failed to create items on the server : "  + e.Message,e);
                    FailedMessageCount += itemBuffer.Count;
                }
                ProcessedMessageCount += itemBuffer.Count;
                itemBuffer.Clear();
                bufferedSize = 0;
            }
            else
            {
                Logger.Debug("No messages to commit");
            }
        }

        private void ImportRunner(RawMessageDescriptor msg)
        {
            var secondAttempt = false;
            var completed = false;
            do
            {
                var start = Environment.TickCount;
                try
                {
                    try
                    {
                        Logger.Debug("Importing " + msg.SourceId + " : " + msg.DestinationFolder + @"\" + msg.Subject + " (" + String.Join(", ", msg.Flags) + ") [" + msg.ItemClass + "]");
                        EmailMessage item = PrepareEWSItem(msg);
                        FolderId parentFolderId = GetCreateFolder(previousFolder);
                        item.Save(parentFolderId);
                        Logger.Debug("Imported message " + msg + " into " + _username + "@" + _hostname + " [" +
                                     (Environment.TickCount - start) + "ms]");
                        SucceededMessageCount++;
                    }
                    catch (ServiceLocalException e)
                    {
                        if (
                            Regex.Match(e.Message,
                                @"The type of the object in the store \(\w+\) does not match that of the local object \(\w+\).")
                                .Success)
                        {
                            // this is an error we can ignore, it just means we sent it as an email message, but its actually a meeting request/response, exchange imports it anyway
                            Logger.Debug("Imported meeting request/response " + msg + " into " + _username + "@" +
                                         _hostname +
                                         " [" +
                                         (Environment.TickCount - start) + "ms]");
                            SucceededMessageCount++;
                        }
                        else
                        {
                            throw e;
                        }
                    }
                    catch (ServiceResponseException e)
                    {
                        if (e.Message.Equals("Operation would change object type, which is not permitted."))
                        {
                            // This error is caused by trying to import a message which is a bounce, nothing we can do, mark as ignored
                            var match = Regex.Match(msg.RawMessage, "Subject:(.*)");
                            Logger.Warn("Ignored bounce message " + msg + " [" +
                                        (match.Success ? match.Groups[0].ToString() : "") + "] for " + _username +
                                        "@" + _hostname + " [" + (Environment.TickCount - start) + "ms]");
                            IgnoredMessageCount++;
                        }
                        else
                        {
                            throw e;
                        }
                    }
                    completed = true;
                }
                catch (Exception e)
                {
                    if (secondAttempt)
                    {
                        var match = Regex.Match(msg.RawMessage, "^Subject:(.*)");
                        Logger.Error(
                            "Failed to import message " + msg + " [" + (msg.Subject) + "] into " + _username + "@" + _hostname + " [" + (Environment.TickCount - start) + "ms]" +
                            " : " + e.Message, e);
                        FailedMessageCount++;
                        secondAttempt = false;
                    }
                    else
                    {
                        var match = Regex.Match(msg.RawMessage, "^Subject:(.*)");
                        Logger.Warn(
                            "Failed on try 1 of 2 to import message " + msg + " [" +
                            msg.Subject +
                            "] into " + _username + "@" + _hostname + "/" + msg.DestinationFolder + " [" + (Environment.TickCount - start) + "ms]" +
                            " : " + e.Message, e);
                        secondAttempt = true;
                    }
                }
            } while (secondAttempt && !completed);
            ProcessedMessageCount++;
        }

        private EmailMessage PrepareEWSItem(RawMessageDescriptor msg)
        {
            var headersString = msg.RawMessage;
            var headerLimit = msg.RawMessage.IndexOf("\n\n");
            if (headerLimit < 0)
            {
                headerLimit = msg.RawMessage.IndexOf("\r\n\r\n");
            }
            if (headerLimit > 0)
            {
                headersString = msg.RawMessage.Substring(0, headerLimit);
            }
            // sometimes there is no received header, we will have to modify this so Outlook sorts this correctly
            if (!headersString.Contains("Received: ") && !msg.DestinationFolder.Equals("Sent Items"))
            {
                var match = Regex.Match(headersString, @"Date: (.*)");
                if (match.Success)
                {
                    Logger.Debug("Missing Received header, adding dummy header");
                    msg.RawMessage = "Received: from zinkuba.export (127.0.0.1) by\r\n" +
                                     " zinkuba.import (127.0.0.1) with Zinkuba id 0.0.0.0;\r\n" +
                                     " " + match.Groups[1] +
                                     "\r\n" +
                                     msg.RawMessage;
                }
            }

            EmailMessage item = new EmailMessage(service)
            {
                MimeContent = new MimeContent("UTF-8", Encoding.UTF8.GetBytes(msg.RawMessage))
            };

            String folder = msg.DestinationFolder;

            // This is required to be set one way or another, otherwise the message is marked as new (not delivered)
            item.IsRead = !msg.Flags.Contains(MessageFlags.Unread);
            // Set the defaults to no receipts, to be corrected later by flags.
            item.IsReadReceiptRequested = false;
            item.IsDeliveryReceiptRequested = false;
            if (!String.IsNullOrEmpty(msg.ItemClass))
            {
                item.ItemClass = msg.ItemClass;
            }
            else
            {
                // default Item.Class
                item.ItemClass = "IPM.Note";
                // we need to detect the item class
                if (headersString.Contains("Return-Path: <>"))
                {
                    // Looks like it may be a bounce
                    if (headersString.Contains("X-MS-Exchange-Message-Is-Ndr:")
                        || headersString.Contains("Content-Type: multipart/report; report-type=delivery-status;")
                        || headersString.Contains("X-Failed-Recipients:")
                        )
                    {
                        item.ItemClass = "REPORT.IPM.NOTE.NDR";
                    }
                }
            }

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

            return item;
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
                    Folder folder = new Folder(service) { DisplayName = destinationFolderName };
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
                    CommitBufferedMessages();
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
