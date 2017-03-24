using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
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
        private const int MaxBufferSize = 5242880; // 5MB of mails before we force a commit
        private const int MaxBufferMails = 50; // 50 mails before we force a commit
        private const int MaxRetryOnUnknownError = 5;
        private static readonly ILog Logger = LogManager.GetLogger(typeof(ExchangeTarget));

        private readonly string _hostname;
        private readonly string _username;
        private readonly string _password;
        private PCQueue<RawMessageDescriptor, ImportState> _queue;
        private ImportState _lastState;
        private ExchangeService service;
        private List<ExchangeFolder> folders;
        private List<MailFolder> _mainFolderList = new List<MailFolder>(); 
        private readonly object _folderCreationLock = new object();
        private readonly SmartThreadPool pool;
        private String previousFolder;
        private List<ExchangeItemContainer> itemBuffer;
        private int bufferedSize;
        private bool _createAllFolders;


        public ExchangeTarget(string hostname, string username, string password, bool createAllFolders)
        {
            _createAllFolders = createAllFolders;
            _hostname = hostname;
            _username = username;
            _password = password;
            itemBuffer = new List<ExchangeItemContainer>();
            pool = new SmartThreadPool() { MaxThreads = 5 };
        }

        public void Process(RawMessageDescriptor message)
        {
            if (Status == MessageProcessorStatus.Started || Status == MessageProcessorStatus.Initialised)
            {
                Status = MessageProcessorStatus.Started;
                // limit the size of the queue
                while (_queue.Count > 150)
                {
                    Thread.Sleep(1000);
                }
                _queue.Consume(message);
            }
            else
            {
                FailedMessageCount++;
                throw new Exception("Cannot process, exchange target is not started.");
            }

        }

        public override void Initialise(List<MailFolder> folderList)
        {
            Status = MessageProcessorStatus.Initialising;
            _mainFolderList = folderList;
            service = ExchangeHelper.ExchangeConnect(_hostname, _username, _password);
            //service.TraceEnabled = true;
            //service.TraceFlags = TraceFlags.All;
            folders = new List<ExchangeFolder>();
            ExchangeHelper.GetAllSubFolders(service, new ExchangeFolder { Folder = Folder.Bind(service, WellKnownFolderName.MsgFolderRoot) }, folders, false);
            if (_createAllFolders)
            {
                // we need to create all folders which don't exist
                foreach (var mailFolder in _mainFolderList)
                {
                    GetCreateFolder(mailFolder.DestinationFolder);
                }
            }
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
            // Commit messages if new folder or the buffer is maxed out (either size wise or count wise)
            if ((!String.IsNullOrEmpty(previousFolder) && !previousFolder.Equals(msg.DestinationFolder)) || bufferedSize >= MaxBufferSize || itemBuffer.Count >= MaxBufferMails)
            {
                CommitBufferedMessages();
            }
            try
            {
                Logger.Debug("Buffering " + msg.SourceId + " : " + msg.DestinationFolder + @"\" + msg.Subject + " (" + String.Join(", ", msg.Flags) + ") [" + msg.ItemClass + "]");
                EmailMessage item = PrepareEWSItem(msg);
                itemBuffer.Add(new ExchangeItemContainer() {ExchangeItem = item, MsgDescriptor = msg});
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
            int tryCount = 0;
            var retryItems = new List<ExchangeItemContainer>();
            while (itemBuffer.Count > 0)
            {
                Logger.Debug("Committing " + itemBuffer.Count + " messages to EWS.");
                retryItems.Clear();
                tryCount++;
                var start = Environment.TickCount;
                try
                {
                    try
                    {
                        FolderId parentFolderId = GetCreateFolder(previousFolder);
                        var response =
                            service.CreateItems(
                                itemBuffer.Select(exchangeItemContainer => exchangeItemContainer.ExchangeItem)
                                    .AsEnumerable(), parentFolderId,
                                MessageDisposition.SaveOnly, null);
                        if (response.OverallResult != ServiceResult.Success)
                        {
                            Logger.Warn("Looks like some items succeeded, others failed, checking.");
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
                                    if (serviceResponse.ErrorCode == ServiceError.ErrorTimeoutExpired ||
                                        serviceResponse.ErrorCode == ServiceError.ErrorBatchProcessingStopped ||
                                        serviceResponse.ErrorCode == ServiceError.ErrorServerBusy)
                                    {
                                        // we can attempt these again
                                        Logger.Warn("Failed to import message " +
                                                    itemBuffer[count].MsgDescriptor.Subject + "[" +
                                                    itemBuffer[count].ExchangeItem.ItemClass + "] into " + _username +
                                                    "@" +
                                                    _hostname +
                                                    "/" +
                                                    previousFolder + " [" + (Environment.TickCount - start) + "ms]" +
                                                    ", will retry, not a permanent error : [" +
                                                    serviceResponse.ErrorCode + "]" +
                                                    serviceResponse.ErrorMessage);
                                        retryItems.Add(itemBuffer[count]);
                                    }
                                    else
                                    {
                                        Logger.Error("Failed to import message " +
                                                     itemBuffer[count].MsgDescriptor.Subject + "[" +
                                                     itemBuffer[count].ExchangeItem.ItemClass + "] into " + _username +
                                                     "@" +
                                                     _hostname +
                                                     "/" +
                                                     previousFolder + " [" + (Environment.TickCount - start) + "ms]" +
                                                     ", permanent error : [" + serviceResponse.ErrorCode + "]" +
                                                     serviceResponse.ErrorMessage);
                                        FailedMessageCount++;
                                    }
                                }
                                count++;
                            }
                            Logger.Warn("Saved " + successCount + " of a possible " + itemBuffer.Count +
                                        " messages into " +
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
                        if (
                            Regex.Match(e.Message,
                                @"The type of the object in the store \(\w+\) does not match that of the local object \(\w+\).")
                                .Success)
                        {
                            // this is an error we can ignore, it just means we sent it as an email message, but its actually a meeting request/response, exchange imports it anyway
                            Logger.Info("Saved " + itemBuffer.Count + " messages [" + (bufferedSize/1024) + "Kb] into " +
                                        _username + "@" + _hostname + "/" + previousFolder + " [" +
                                        (Environment.TickCount - start) + "ms]");
                            SucceededMessageCount += itemBuffer.Count;
                        }
                        else
                        {
                            Logger.Error("Got a service local exception I didn't understand.");
                            // this will be caught by the bigger try and processed correctly
                            throw e;
                        }
                    }
                    catch (ServerBusyException e)
                    {
                        var retryWait = 30000;
                        // make sure its reasonable
                        if (e.BackOffMilliseconds > retryWait && e.BackOffMilliseconds < 300000)
                        {
                            retryWait = e.BackOffMilliseconds;
                        }
                        Logger.Error("Failed to create items, server too busy, backing off (" + (int) (retryWait/1000) +
                                     "s) and trying again : " +
                                     e.Message);
                        // Lets loop back around and try again.
                        Thread.Sleep(retryWait); // Sleep x seconds
                        continue; // run loop again
                    }
                    catch (ServiceRequestException e)
                    {
                        if (e.GetBaseException() is System.Net.WebException)
                        {
                            var retryWait = 30000;
                            Logger.Error("Failed to create items, error connecting to server, backing off (" + (int)(retryWait / 1000) +"s) and trying again : " + e.Message);
                            // Lets loop back around and try again.
                            Thread.Sleep(retryWait); // Sleep x seconds
                            continue; // run loop again
                        }
                        else
                        {
                            throw e;
                        }
                    }
                }
                catch (Exception e)
                {
                    // This exception means something happened on the server, we try again MaxRetryOnUnknownError times, then fail
                    if (tryCount > MaxRetryOnUnknownError)
                    {
                        Logger.Error("Failed to create items on the server : " + e.Message, e);
                        foreach (var item in itemBuffer)
                        {
                            Logger.Error("Gave up on inserting [" + item.MsgDescriptor.Subject + "] into " + _username + "@" + _hostname + "/" +
                                         previousFolder);
                        }
                        FailedMessageCount += itemBuffer.Count;
                    }
                    else
                    {
                        Logger.Warn("Failed to create items on server, trying " + (MaxRetryOnUnknownError - tryCount) + " more times : " +
                                    e.Message);
                        Thread.Sleep(30000); // sleep 30 second
                        continue; // run loop again
                    }
                }
                ProcessedMessageCount += itemBuffer.Count - retryItems.Count;
                itemBuffer.Clear();
                bufferedSize = 0;
                // if there are retry items, we need to put them back in the buffer
                if (retryItems.Count > 0)
                {
                    Logger.Info("Will retry " + retryItems.Count + " items, will wait for 30s before retrying.");
                    itemBuffer.AddRange(retryItems);
                    bufferedSize = retryItems.Sum(m => m.MsgDescriptor.RawMessage.Length);
                    Thread.Sleep(30000); // wait 30 seconds before retrying
                }
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
            if (msg.DestinationFolder.Equals("Sent Items"))
            {
                // we need to set a sent date property otherwise dates don't show up properly in exchange
                DateTime? sentDate = msg.SentDateTime;
                if (sentDate == null)
                {
                    var match = Regex.Match(headersString, @"Date: (.*)");
                    if (match.Success)
                    {
                        try
                        {
                            sentDate = DateTime.Parse(match.Groups[1].Value);
                        }
                        catch (Exception e)
                        {
                            Logger.Error("Failed to parse header date " + match.Groups[1] + " to a date for sending.");
                        }
                    }
                }
                if (sentDate == null)
                {
                    Logger.Error("Failed to set sent date on " + msg.Subject);
                }
                else
                {
                    try
                    {
                        item.SetExtendedProperty(ExchangeHelper.MsgPropertyDateTimeSent,
                            sentDate.Value.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ"));
                        item.SetExtendedProperty(ExchangeHelper.MsgPropertyDateTimeReceived,
                            sentDate.Value.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ"));
                    }
                    catch (Exception e)
                    {
                        Logger.Error("Failed to set sent item date on " + msg.Subject + " to " + sentDate, e);
                    }
                }
            }

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
                    Logger.Warn("Failed to set flag on " + folder + @"\" + msg.Subject + ", ignoring flag.", e);
                }
            }

            if (msg.FlagIcon != FlagIcon.None)
            {
                item.SetExtendedProperty(ExchangeHelper.PidTagFlagStatus, 2);
                item.SetExtendedProperty(ExchangeHelper.PidTagFollowupIcon, ExchangeHelper.ConvertFlagIcon(msg.FlagIcon));
            }

            if (service.RequestedServerVersion == ExchangeVersion.Exchange2013 && msg.FollowUpFlag != null)
            {
                item.Flag = new Flag()
                {
                    DueDate = msg.FollowUpFlag.DueDateTime,
                    StartDate = msg.FollowUpFlag.StartDateTime,
                    CompleteDate = msg.FollowUpFlag.CompleteDateTime,
                    FlagStatus = ExchangeHelper.ConvertFlagStatus(msg.FollowUpFlag.Status),
                };
            }

            try
            {
                if (!msg.Flags.Contains(MessageFlags.Draft))
                {
                    item.SetExtendedProperty(ExchangeHelper.MsgFlagRead, 1);
                }
                if (msg.Importance != null) item.Importance = (Importance)msg.Importance;
                if (msg.Sensitivity != null) item.Sensitivity = (Sensitivity)msg.Sensitivity;
                if (msg.ReminderDueBy != null) item.ReminderDueBy = (DateTime)msg.ReminderDueBy;
                item.Categories = new StringList(msg.Categories);
            }
            catch (Exception e)
            {
                Logger.Warn(
                    "Failed to set metadata on " + folder + @"\" + msg.Subject + ", ignoring metadata.", e);
            }

            return item;
        }

        private FolderId GetCreateFolder(string destinationFolder, bool secondAttempt = false)
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
                    try
                    {
                        folder.Save(parentFolderId);
                    }
                    catch (Exception e)
                    {
                        // If the folder exists, we need to refresh and have another crack
                        if (e.Message.Equals("A folder with the specified name already exists.") && !secondAttempt)
                        {
                            Logger.Warn("Looks like the folder "+ destinationFolder + " was created under our feet, refreshing the folder list. We will only attempt this once per folder.");
                            // Oops, the folders have been updated on the server, we need a refresh
                            var newFolders = new List<ExchangeFolder>();
                            ExchangeHelper.GetAllSubFolders(service, new ExchangeFolder { Folder = Folder.Bind(service, WellKnownFolderName.MsgFolderRoot) }, newFolders, false);
                            folders = newFolders;
                            // lets try again
                            return GetCreateFolder(destinationFolder, true);
                        }
                        throw e;
                    }
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
                Logger.Info("Closing Exchange Target");
            }
        }

        public List<string> ImportedIds { get; private set; }
    }

    public class ExchangeItemContainer
    {
        public Item ExchangeItem;
        public RawMessageDescriptor MsgDescriptor;
    }
}
