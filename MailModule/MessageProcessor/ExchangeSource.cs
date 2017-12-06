using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO.Ports;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Xml;
using log4net;
using Microsoft.Exchange.WebServices.Data;
using Microsoft.Office.Interop.Outlook;
using Zinkuba.MailModule.API;
using Zinkuba.MailModule.MessageDescriptor;
using Exception = System.Exception;
using Folder = Microsoft.Exchange.WebServices.Data.Folder;

namespace Zinkuba.MailModule.MessageProcessor
{
    public class ExchangeSource : BaseMessageProcessor, IMessageWriter<RawMessageDescriptor>, IMessageSource
    {
        private static readonly ILog Logger = LogManager.GetLogger(typeof(ExchangeSource));

        private readonly string _username;
        private readonly string _password;
        private readonly string _hostname;
        private readonly DateTime _startDate;
        private readonly DateTime _endDate;
        private readonly List<string> _limitFolderList;
        private readonly List<MailFolder> _mainFolderList = new List<MailFolder>(); 
        private Thread _sourceThread;
        private ExchangeService service;
        private List<ExchangeFolder> folders;
        private int _totalMessages;
        private int _pageSize = 20;

        public IMessageReader<RawMessageDescriptor> NextReader { get; set; }

        public int TotalMessages
        {
            get { return _totalMessages; }
            private set { _totalMessages = value; OnTotalMessagesChanged(); }
        }

        public bool TestOnly { get; set; }
        public bool IncludePublicFolders { get; set; }

        public event EventHandler TotalMessagesChanged;

        protected virtual void OnTotalMessagesChanged()
        {
            EventHandler handler = TotalMessagesChanged;
            if (handler != null) handler(this, EventArgs.Empty);
        }

        public ExchangeSource(String username, String password, String hostname, DateTime startDate, DateTime endDate, List<String> limitFolderList)
        {
            _username = username;
            _password = password;
            _hostname = hostname;
            _startDate = startDate;
            _endDate = endDate;
            _limitFolderList = (limitFolderList == null || limitFolderList.Count == 0) ? null : limitFolderList;
            Name = username;
        }

        public override void Initialise(List<MailFolder> folderList)
        {
            Status = MessageProcessorStatus.Initialising;
            service = ExchangeHelper.ExchangeConnect(_hostname, _username, _password);
            folders = new List<ExchangeFolder>();
            // Use Exchange Helper to get all the folders for this account
            ExchangeHelper.GetAllSubFolders(service,
                new ExchangeFolder() { Folder = Folder.Bind(service, WellKnownFolderName.MsgFolderRoot) }, folders,
                false);
            if (IncludePublicFolders)
            {
                Logger.Debug("Including Public Folders");
                ExchangeHelper.GetAllSubFolders(service,
                    new ExchangeFolder() { Folder = Folder.Bind(service, WellKnownFolderName.PublicFoldersRoot), IsPublicFolder = true }, folders,
                    false);
            }

            // Are we limited folders to a specific list?
            if (_limitFolderList != null)
            {
                var newFolders = new List<ExchangeFolder>();
                foreach (var mailbox in _limitFolderList)
                {
                    var mailboxMatch = mailbox.ToLower().Replace('/', '\\'); ;
                    newFolders.AddRange(folders.Where(folder => folder.FolderPath.ToLower().Equals(mailboxMatch)));
                }
                folders = newFolders;
            }

            // Scan the folders to get message counts
            ExchangeHelper.GetFolderSummary(service, folders, _startDate, _endDate);
            folders.ForEach(folder => TotalMessages += !TestOnly ? folder.MessageCount : (folder.MessageCount > 20 ? 20 : folder.MessageCount));
            Logger.Debug("Found " + folders.Count + " folders and " + TotalMessages + " messages.");

            // Now build the folder list that we pass on to the next folders.
            foreach (var exchangeFolder in folders)
            {
                var folder = new MailFolder()
                {
                    SourceFolder = exchangeFolder.FolderPath,
                    DestinationFolder = exchangeFolder.MappedDestination,
                    MessageCount = exchangeFolder.MessageCount,
                };
                _mainFolderList.Add(folder);
            }
            // Now initialise the next read, I am not going to start reading unless I know the pipeline is groovy
            NextReader.Initialise(_mainFolderList);
            Status = MessageProcessorStatus.Initialised;
            Logger.Info("ExchangeExporter Initialised");
        }


        public Type OutMessageDescriptorType()
        {
            return typeof(RawMessageDescriptor);
        }

        public void Start()
        {
            _sourceThread = new Thread(Run) { IsBackground = true, Name = "exchangeExporter-" + _username };
            _sourceThread.Start();
        }

        private void Run()
        {
            try
            {
                Status = MessageProcessorStatus.Started;
                var fullPropertySet = new PropertySet(PropertySet.FirstClassProperties)
                {
                    EmailMessageSchema.IsRead,
                    EmailMessageSchema.IsReadReceiptRequested,
                    EmailMessageSchema.IsDeliveryReceiptRequested,
                    ItemSchema.DateTimeSent,
                    ItemSchema.DateTimeReceived,
                    ItemSchema.DateTimeCreated,
                    ItemSchema.ItemClass,
                    ItemSchema.MimeContent,
                    ItemSchema.Categories,
                    ItemSchema.Importance,
                    ItemSchema.InReplyTo,
                    ItemSchema.IsFromMe,
                    ItemSchema.IsReminderSet,
                    ItemSchema.IsResend,
                    ItemSchema.IsDraft,
                    ItemSchema.ReminderDueBy,
                    ItemSchema.Sensitivity,
                    ItemSchema.Subject,
                    ItemSchema.Id,
                    ExchangeHelper.MsgPropertyContentType,
                    ExchangeHelper.PidTagFollowupIcon,
                };
                if (service.RequestedServerVersion != ExchangeVersion.Exchange2007_SP1)
                {
                    fullPropertySet.Add(ItemSchema.ConversationId);
                    fullPropertySet.Add(ItemSchema.IsAssociated);
                }
                if (service.RequestedServerVersion == ExchangeVersion.Exchange2013)
                {
                    fullPropertySet.Add(ItemSchema.ArchiveTag);
                    fullPropertySet.Add(ItemSchema.Flag);
                    fullPropertySet.Add(ItemSchema.IconIndex);
                }
                SearchFilter.SearchFilterCollection filter = new SearchFilter.SearchFilterCollection();
                filter.LogicalOperator = LogicalOperator.And;
                if (_startDate != null)
                {
                    Logger.Debug("Getting mails from " + _startDate);
                    filter.Add(new SearchFilter.IsGreaterThanOrEqualTo(ItemSchema.DateTimeReceived, _startDate));
                }
                if (_endDate != null)
                {
                    Logger.Debug("Getting mails up until " + _endDate);
                    filter.Add(new SearchFilter.IsLessThanOrEqualTo(ItemSchema.DateTimeReceived, _endDate));
                }
                foreach (var exchangeFolder in folders)
                {
                    ItemView view = new ItemView(_pageSize, 0, OffsetBasePoint.Beginning);
                    view.PropertySet = PropertySet.IdOnly;
                    List<EmailMessage> emails = new List<EmailMessage>();
                    Boolean more = true;
                    while (more)
                    {
                        try
                        {
                            more = FindExchangeMessages(exchangeFolder, filter, view, emails, fullPropertySet);
                            if (emails.Count > 0)
                            {
                                try
                                {
                                    foreach (var emailMessage in emails)
                                    {
                                        try
                                        {
                                            String subject;
                                            if (!emailMessage.TryGetProperty(ItemSchema.Subject, out subject) ||
                                                subject == null)
                                            {
                                                Logger.Warn("Item " + emailMessage.Id.UniqueId + " has no subject assigned, unable to determine subject.");
                                            }
                                            Logger.Debug("Exporting " + emailMessage.Id.UniqueId + " from " + exchangeFolder.FolderPath + " : " + subject);
                                            var flags = new Collection<MessageFlags>();
                                            Boolean flag;
                                            if (emailMessage.TryGetProperty(EmailMessageSchema.IsRead, out flag) &&
                                                !flag) flags.Add(MessageFlags.Unread);
                                            if (emailMessage.TryGetProperty(ItemSchema.IsDraft, out flag) && flag)
                                                flags.Add(MessageFlags.Draft);
                                            if (
                                                emailMessage.TryGetProperty(EmailMessageSchema.IsReadReceiptRequested,
                                                    out flag) && flag) flags.Add(MessageFlags.ReadReceiptRequested);
                                            if (
                                                emailMessage.TryGetProperty(
                                                    EmailMessageSchema.IsDeliveryReceiptRequested, out flag) && flag)
                                                flags.Add(MessageFlags.DeliveryReceiptRequested);
                                            if (emailMessage.TryGetProperty(ItemSchema.IsReminderSet, out flag) && flag)
                                                flags.Add(MessageFlags.ReminderSet);
                                            if (emailMessage.TryGetProperty(ItemSchema.IsAssociated, out flag) && flag)
                                                flags.Add(MessageFlags.Associated);
                                            if (emailMessage.TryGetProperty(ItemSchema.IsFromMe, out flag) && flag)
                                                flags.Add(MessageFlags.FromMe);
                                            if (emailMessage.TryGetProperty(ItemSchema.IsResend, out flag) && flag)
                                                flags.Add(MessageFlags.Resend);
                                            var message = new RawMessageDescriptor
                                            {
                                                SourceId = emailMessage.Id.UniqueId,
                                                Subject = subject,
                                                Flags = flags,
                                                RawMessage = "",
                                                SourceFolder = exchangeFolder.FolderPath,
                                                DestinationFolder = exchangeFolder.MappedDestination,
                                                IsPublicFolder = exchangeFolder.IsPublicFolder,
                                            };
                                            Object result;
                                            if (emailMessage.TryGetProperty(ItemSchema.MimeContent, out result) &&
                                                result != null)
                                                message.RawMessage = Encoding.UTF8.GetString(emailMessage.MimeContent.Content);
                                            if (emailMessage.TryGetProperty(ItemSchema.ItemClass, out result) &&
                                                result != null) message.ItemClass = emailMessage.ItemClass;
                                            if (emailMessage.TryGetProperty(ItemSchema.IconIndex, out result) &&
                                                result != null) message.IconIndex = (int)emailMessage.IconIndex;
                                            if (emailMessage.TryGetProperty(ItemSchema.Importance, out result) &&
                                                result != null) message.Importance = (int)emailMessage.Importance;
                                            if (emailMessage.TryGetProperty(ItemSchema.Sensitivity, out result) &&
                                                result != null) message.Sensitivity = (int)emailMessage.Sensitivity;
                                            if (emailMessage.TryGetProperty(ItemSchema.InReplyTo, out result) &&
                                                result != null) message.InReplyTo = emailMessage.InReplyTo;
                                            if (emailMessage.TryGetProperty(ItemSchema.ConversationId, out result) &&
                                                result != null)
                                                message.ConversationId = emailMessage.ConversationId.ChangeKey;
                                            if (emailMessage.TryGetProperty(ItemSchema.ReminderDueBy, out result) &&
                                                result != null) message.ReminderDueBy = emailMessage.ReminderDueBy;
                                            if (
                                                emailMessage.TryGetProperty(ExchangeHelper.PidTagFollowupIcon,
                                                    out result) && result != null)
                                                message.FlagIcon = ExchangeHelper.ConvertFlagIcon((int)result);
                                            if (emailMessage.TryGetProperty(ItemSchema.DateTimeReceived, out result) &&
                                                result != null)
                                                message.ReceivedDateTime = emailMessage.DateTimeReceived;
                                            if (emailMessage.TryGetProperty(ItemSchema.DateTimeSent, out result) &&
                                                result != null) message.SentDateTime = emailMessage.DateTimeSent;
                                            if (emailMessage.TryGetProperty(ItemSchema.Flag, out result) &&
                                                result != null)
                                                message.FollowUpFlag = new FollowUpFlag()
                                                {
                                                    StartDateTime = ((Flag)result).StartDate,
                                                    DueDateTime = ((Flag)result).DueDate,
                                                    CompleteDateTime = ((Flag)result).CompleteDate,
                                                    Status =
                                                        ExchangeHelper.ConvertFlagStatus(((Flag)result).FlagStatus),
                                                };
                                            if (emailMessage.TryGetProperty(ItemSchema.Categories, out result) &&
                                                result != null && emailMessage.Categories.Count > 0)
                                            {
                                                foreach (var category in emailMessage.Categories)
                                                {
                                                    message.Categories.Add(category);
                                                }
                                            }
                                            if (emailMessage.ExtendedProperties != null)
                                            {
                                                foreach (var extendedProperty in emailMessage.ExtendedProperties)
                                                {
                                                    if (
                                                        extendedProperty.PropertyDefinition.Equals(
                                                            ExchangeHelper.MsgPropertyContentType))
                                                    {
                                                        if (extendedProperty.Value.ToString().Contains("signed-data"))
                                                        {
                                                            message.IsEncrypted = true;
                                                        }
                                                    }
                                                }
                                            }
                                            NextReader.Process(message);
                                            SucceededMessageCount++;
                                        }
                                        catch (Exception e)
                                        {
                                            Logger.Error("Failed to load properties for message " + emailMessage.Id.UniqueId, e);
                                            FailedMessageCount++;
                                        }
                                    }
                                }
                                catch (Exception e)
                                {
                                    Logger.Error("Failed to load properties for messages in " + exchangeFolder.FolderPath, e);
                                    FailedMessageCount += emails.Count;
                                }
                                ProcessedMessageCount += emails.Count;
                            }
                            if (more)
                            {
                                view.Offset += _pageSize;
                            }
                        }
                        catch (Exception e)
                        {
                            Logger.Error("Failed to find results against folder " + exchangeFolder.FolderPath, e);
                            more = false;
                        }
                    }
                }
            }
            catch (Exception e)
            {
                Logger.Error("Failed to run exporter", e);

            }
            finally
            {
                Close();
            }
        }

        private bool FindExchangeMessages(ExchangeFolder exchangeFolder, SearchFilter.SearchFilterCollection filter, ItemView view, List<EmailMessage> emails, PropertySet fullPropertySet)
        {
            int majorFailureCount = 0;
            int tryCount = 0;
            do
            {
                emails.Clear();
                tryCount++;
                var start = Environment.TickCount;
                try
                {
                    Logger.Debug("Retrieving next email messages from " + exchangeFolder.FolderPath);
                    // Get results
                    FindItemsResults<Item> findResults = exchangeFolder.Folder.FindItems(filter, view);
                    // Filter out non mail items
                    foreach (var item in findResults)
                    {
                        if (item is EmailMessage)
                        {
                            try
                            {
                                emails.Add((EmailMessage) item);
                                if (TestOnly && emails.Count > 20)
                                {
                                    break;
                                }
                            }
                            catch (Exception e)
                            {
                                FailedMessageCount++;
                                Logger.Error("Failed to extract email message from results, skipping : " + e.Message, e);
                            }
                        }
                        else
                        {
                            Logger.Warn("Will not export message, item is not a mail message [" + item.GetType() + "]");
                            IgnoredMessageCount++;
                        }
                    }
                    if (emails.Count > 0)
                    {
                        // Load properties for our final list of mails
                        Logger.Debug("Retrieving " + emails.Count + " emails properties from " +
                                     exchangeFolder.FolderPath);
                        service.LoadPropertiesForItems(emails, fullPropertySet);
                    }
                    Logger.Debug("Retrieved " + emails.Count + " messages from " + exchangeFolder.FolderPath + " in " + (Environment.TickCount - start) + "ms, trying " + tryCount + " times. " + (findResults.MoreAvailable ? "More" : "No more") + " messages in folder available.");
                    return findResults.MoreAvailable && !TestOnly;
                }
                catch (ServerBusyException e)
                {
                    Logger.Warn(
                        "Received Server Busy Response (ServerBusyException), will back off and try again after (30s) : " +
                        e.Message);
                    Thread.Sleep(30000);
                }
                catch (XmlException e)
                {
                    Logger.Error("XML Response invalid (XmlException), will attempt again : " +
                                e.Message);
                    Thread.Sleep(1000);
                }
                catch (TimeoutException e)
                {
                    Logger.Error("Request timed out with TimeoutException, will back off and try again after (30s) : " +
                                e.Message);
                    Thread.Sleep(30000);
                }
                catch (ServiceRequestException e)
                {
                    Logger.Error(
                        "Service request failed with ServiceRequestException, this may not be a permanent error, will back off and try again after (30s) : " +
                        e.Message);
             
                Thread.Sleep(30000);
                } catch (Exception e)
                {
                    majorFailureCount++;
                    if (majorFailureCount >= 10)
                    {
                        throw e;
                    }
                    else
                    {
                        Logger.Error("Failed to get emails/properties, will attempt " + (10 - majorFailureCount) + " more times after 30s pause.",e);
                        Thread.Sleep(30000);
                    }
                }
            } while(true);
        }


        public override void Close()
        {
            if (!Closed)
            {
                Closed = true;
                NextReader.Close();
            }
        }

        /*
        public void GetAllFolders(ExchangeService service, List<Folder> completeListOfFolderIds)
        {
            FolderView folderView = new FolderView(int.MaxValue);
            FindFoldersResults findFolderResults = service.FindFolders(WellKnownFolderName.PublicFoldersRoot, folderView);
            foreach (Folder folder in findFolderResults)
            {
                completeListOfFolderIds.Add(folder);
                FindAllSubFolders(service, folder.Id, completeListOfFolderIds);
            }
        }

        private void FindAllSubFolders(ExchangeService service, FolderId parentFolderId, List<Folder> completeListOfFolderIds)
        {
            //search for sub folders
            FolderView folderView = new FolderView(int.MaxValue);
            FindFoldersResults foundFolders = service.FindFolders(parentFolderId, folderView);

            // Add the list to the growing complete list
            completeListOfFolderIds.AddRange(foundFolders);

            // Now recurse
            foreach (Folder folder in foundFolders)
            {
                FindAllSubFolders(service, folder.Id, completeListOfFolderIds);
            }
        }
         */

    }

    internal class ExchangeFolder
    {
        public int MessageCount;
        public FolderId FolderId;
        public String FolderPath;
        public Folder Folder;
        public string MappedDestination { get; set; }
        public bool IsPublicFolder;
    }
}
