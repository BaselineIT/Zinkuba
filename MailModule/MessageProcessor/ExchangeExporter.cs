using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO.Ports;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using log4net;
using Microsoft.Exchange.WebServices.Data;
using Zinkuba.MailModule.API;
using Zinkuba.MailModule.MessageDescriptor;
using Exception = System.Exception;
using Folder = Microsoft.Exchange.WebServices.Data.Folder;

namespace Zinkuba.MailModule.MessageProcessor
{
    public class ExchangeExporter : BaseMessageProcessor, IMessageWriter<RawMessageDescriptor>, IMessageSource
    {
        private static readonly ILog Logger = LogManager.GetLogger(typeof(ExchangeExporter));

        private readonly string _username;
        private readonly string _password;
        private readonly string _hostname;
        private readonly DateTime _startDate;
        private readonly DateTime _endDate;
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
        public event EventHandler TotalMessagesChanged;

        protected virtual void OnTotalMessagesChanged()
        {
            EventHandler handler = TotalMessagesChanged;
            if (handler != null) handler(this, EventArgs.Empty);
        }

        public ExchangeExporter(String username, String password, String hostname, DateTime startDate, DateTime endDate)
        {
            _username = username;
            _password = password;
            _hostname = hostname;
            _startDate = startDate;
            _endDate = endDate;
            Name = username;
        }

        public override void Initialise()
        {
            Status = MessageProcessorStatus.Initialising;
            NextReader.Initialise();
            service = ExchangeHelper.ExchangeConnect(_hostname, _username, _password);
            folders = new List<ExchangeFolder>();
            ExchangeHelper.GetAllFolders(service, new ExchangeFolder() { Folder = Folder.Bind(service, WellKnownFolderName.MsgFolderRoot) }, folders, false);
            ExchangeHelper.GetFolderSummary(service, folders, _startDate, _endDate);
            folders.ForEach(folder => TotalMessages += !TestOnly ? folder.MessageCount : (folder.MessageCount > 20 ? 20 : folder.MessageCount));
            Logger.Debug("Found " + folders.Count + " folders and " + TotalMessages + " messages.");
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
                var contentTypeProperty = new ExtendedPropertyDefinition(DefaultExtendedPropertySet.InternetHeaders, "Content-Type", MapiPropertyType.String);
                var fullPropertySet = new PropertySet(PropertySet.FirstClassProperties)
                {
                    EmailMessageSchema.IsRead,
                    EmailMessageSchema.IsReadReceiptRequested,
                    EmailMessageSchema.IsDeliveryReceiptRequested,
                    EmailMessageSchema.Id,
                    ItemSchema.ItemClass,
                    ItemSchema.Subject,
                    ItemSchema.MimeContent,
                    ItemSchema.Categories,
                    ItemSchema.ConversationId,
                    ItemSchema.Importance,
                    ItemSchema.InReplyTo,
                    ItemSchema.IsAssociated,
                    ItemSchema.IsFromMe,
                    ItemSchema.IsReminderSet,
                    ItemSchema.IsResend,
                    ItemSchema.IsDraft,
                    ItemSchema.ReminderDueBy,
                    ItemSchema.Sensitivity,
                    ItemSchema.Subject,
                    ItemSchema.Id,
                    contentTypeProperty,
                };
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
                        emails.Clear();
                        try
                        {
                            FindItemsResults<Item> findResults = null;
                            bool firstAttempt = true;
                            bool success = false;
                            do
                            {
                                try
                                {
                                    findResults = exchangeFolder.Folder.FindItems(filter, view);
                                    success = true;
                                }
                                catch (Exception ex)
                                {
                                    // this is normally a timeout exception, so lets try again
                                    if (firstAttempt)
                                    {
                                        Logger.Warn("Error accessing Exchange Webservice for emails, will attempt again : " +
                                                    ex.Message);
                                        firstAttempt = false;
                                    }
                                    else
                                    {
                                        throw ex;
                                    }
                                }
                            } while (!success);
                            Logger.Debug("Got list of " + findResults.Items.Count + "/" + (findResults.TotalCount - view.Offset) + " messageIds from " + exchangeFolder.FolderPath);
                            foreach (var item in findResults)
                            {
                                if (item is EmailMessage)
                                {
                                    try
                                    {
                                        emails.Add((EmailMessage)item);
                                        if (TestOnly && emails.Count > 20)
                                        {
                                            break;
                                        }
                                    }
                                    catch (Exception e)
                                    {
                                        FailedMessageCount++;
                                        Logger.Error("Failed to extract email message from results : " + e.Message, e);
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
                                try
                                {
                                    Logger.Debug("Retrieving " + emails.Count + " emails properties from " + exchangeFolder.FolderPath);
                                    firstAttempt = true;
                                    success = false;
                                    do
                                    {
                                        try
                                        {
                                            service.LoadPropertiesForItems(emails, fullPropertySet);
                                            success = true;
                                        }
                                        catch (Exception ex)
                                        {
                                            // this is normally a timeout exception, so lets try again
                                            if (firstAttempt)
                                            {
                                                Logger.Warn("Error accessing Exchange Webservice for email properties, will attempt again : " +
                                                            ex.Message);
                                                firstAttempt = false;
                                            }
                                            else
                                            {
                                                throw ex;
                                            }
                                        }
                                    } while (!success);
                                    foreach (var emailMessage in emails)
                                    {
                                        Logger.Debug("Exporting " + emailMessage.Id.UniqueId + " from " + exchangeFolder.FolderPath + " : " + emailMessage.Subject);
                                        var flags = new Collection<MessageFlags>();
                                        Boolean flag;
                                        if (emailMessage.TryGetProperty(EmailMessageSchema.IsRead, out flag) && !flag) flags.Add(MessageFlags.Unread);
                                        if (emailMessage.TryGetProperty(ItemSchema.IsDraft, out flag) && flag) flags.Add(MessageFlags.Draft);
                                        if (emailMessage.TryGetProperty(EmailMessageSchema.IsReadReceiptRequested, out flag) && flag) flags.Add(MessageFlags.ReadReceiptRequested);
                                        if (emailMessage.TryGetProperty(EmailMessageSchema.IsDeliveryReceiptRequested, out flag) && flag) flags.Add(MessageFlags.DeliveryReceiptRequested);
                                        if (emailMessage.TryGetProperty(ItemSchema.IsReminderSet, out flag) && flag) flags.Add(MessageFlags.ReminderSet);
                                        if (emailMessage.TryGetProperty(ItemSchema.IsAssociated, out flag) && flag) flags.Add(MessageFlags.Associated);
                                        if (emailMessage.TryGetProperty(ItemSchema.IsFromMe, out flag) && flag) flags.Add(MessageFlags.FromMe);
                                        if (emailMessage.TryGetProperty(ItemSchema.IsResend, out flag) && flag) flags.Add(MessageFlags.Resend);
                                        var message = new RawMessageDescriptor
                                        {
                                            SourceId = emailMessage.Id.UniqueId,
                                            RawMessage = Encoding.UTF8.GetString(emailMessage.MimeContent.Content),
                                            Flags = flags,
                                            SourceFolder = exchangeFolder.FolderPath,
                                            DestinationFolder = exchangeFolder.MappedDestination,
                                            ItemClass = emailMessage.ItemClass
                                        };
                                        Object result;
                                        if (emailMessage.TryGetProperty(ItemSchema.IconIndex, out result) && result != null) message.IconIndex = (int)emailMessage.IconIndex;
                                        if (emailMessage.TryGetProperty(ItemSchema.Importance, out result) && result != null) message.Importance = (int)emailMessage.Importance;
                                        if (emailMessage.TryGetProperty(ItemSchema.Sensitivity, out result) && result != null) message.Sensitivity = (int)emailMessage.Sensitivity;
                                        if (emailMessage.TryGetProperty(ItemSchema.InReplyTo, out result) && result != null) message.InReplyTo = emailMessage.InReplyTo;
                                        if (emailMessage.TryGetProperty(ItemSchema.ConversationId, out result) && result != null) message.ConversationId = emailMessage.ConversationId.ChangeKey;
                                        if (emailMessage.TryGetProperty(ItemSchema.ReminderDueBy, out result) && result != null) message.ReminderDueBy = emailMessage.ReminderDueBy;
                                        if (emailMessage.TryGetProperty(ItemSchema.Categories, out result) && result != null && emailMessage.Categories.Count > 0)
                                        {
                                            foreach (var category in emailMessage.Categories)
                                            {
                                                message.Categories.Add(category);
                                            }
                                        }
                                        if (NextReader.Status == MessageProcessorStatus.Idle)
                                        {
                                            NextReader.Initialise();
                                        }
                                        if (emailMessage.ExtendedProperties != null)
                                        {
                                            foreach (var extendedProperty in emailMessage.ExtendedProperties)
                                            {
                                                if (extendedProperty.PropertyDefinition.Equals(contentTypeProperty))
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
                                }
                                catch (Exception e)
                                {
                                    Logger.Error("Failed to load properties for messages in " + exchangeFolder.FolderPath, e);
                                    FailedMessageCount += emails.Count;
                                }
                                ProcessedMessageCount += emails.Count;
                            }
                            more = findResults.MoreAvailable && !TestOnly;
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
    }
}
