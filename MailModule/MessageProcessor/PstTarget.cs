using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using log4net;
using Microsoft.Office.Interop.Outlook;
using Zinkuba.MailModule.API;
using Zinkuba.MailModule.MessageDescriptor;
using Exception = System.Exception;

namespace Zinkuba.MailModule.MessageProcessor
{
    public class PstTarget : BaseMessageProcessor, IMessageReader<MsgDescriptor>, IMessageDestination
    {
        private readonly string _saveFolder;
        private static readonly ILog Logger = LogManager.GetLogger(typeof(PstTarget));
        //public String PstPath;
        private Application _outlook;
        private Store _pstHandle;
        private Folder _rootFolder;
        private ImportState _lastState;
        //private QueuedMessageReader<MsgDescriptor, ImportState> _queue;
        private PCQueue<MsgDescriptor, ImportState> _queue;
        private Folder _importFolder;
        private MailItem _importIdsMsg;

        private MailItem ImportIdsMsg
        {
            get
            {
                if (_importIdsMsg == null)
                {
                    var searchResults = _importFolder.Items.OfType<MailItem>().Where(m => m.Subject == "ImportIds").Select(m => m);
                    if (searchResults.Any())
                    {
                        Logger.Debug("Found import id store, retrieving");
                        _importIdsMsg = searchResults.First();
                    }
                    else
                    {
                        Logger.Debug("Creating import id store");
                        _importIdsMsg = _importFolder.Items.Add(OlItemType.olMailItem);
                        _importIdsMsg.Subject = "ImportIds";
                    }
                }
                return _importIdsMsg;
            }
        }

        public List<String> ImportedIds
        {
            get
            {
                if (Status == MessageProcessorStatus.Idle)
                {
                    Initialise();
                }
                try
                {
                    var ids = ImportIdsMsg.Body.Split(new char[] { '\n' }).Where(s => !s.StartsWith("#")).ToList();
                    return ids;
                }
                catch (Exception e)
                {
                    Logger.Error("Failed to get importIds", e);
                    throw;
                }
            }
        }

        private String _pstName
        {
            get { return "Zinkuba-" + (String.IsNullOrWhiteSpace(Name) ? "unknown" : Name); }
        }

        public PstTarget(String saveFolder)
        {
            _saveFolder = saveFolder;
        }

        public void Process(MsgDescriptor message)
        {
            if (Status == MessageProcessorStatus.Started || Status == MessageProcessorStatus.Initialised)
            {
                Status = MessageProcessorStatus.Started;
                _queue.Consume(message);
            }
            else
            {
                FailedMessageCount++;
                throw new Exception("Cannot process, pstTarget is not started.");
            }
        }

        public Type InMessageDescriptorType()
        {
            return typeof(MsgDescriptor);
        }

        #region Queue Methods (Start, Process, Stop)

        public override void Initialise()
        {
            Status = MessageProcessorStatus.Initialising;
            _outlook = new Application();
            var pstPath = Regex.Replace(_saveFolder + @"\" + _pstName + ".pst", @"\\+", @"\");
            Logger.Info("Writing to PST " + pstPath);
            _pstHandle = PstHandle(_outlook.GetNamespace("MAPI"), pstPath, _pstName);
            _rootFolder = _pstHandle.GetRootFolder() as Folder;
            _importFolder = GetCreateFolder(_rootFolder.FolderPath + @"\_zinkuba_import", _outlook);
            _lastState = new ImportState();
            _queue = new PCQueue<MsgDescriptor, ImportState>(Name + "-pstTarget")
            {
                ProduceMethod = ProcessMessage,
                InitialiseProducer = () => _lastState,
                ShutdownProducer = ShutdownQueue
            };
            _queue.Start();
            Status = MessageProcessorStatus.Initialised;
            Logger.Info("PstTarget Initialised");
        }

        private ImportState ProcessMessage(MsgDescriptor msg, ImportState importState)
        {
            if (!msg.SourceFolder.Equals(importState.CurrentFolder))
            {
                if (!String.IsNullOrWhiteSpace(importState.CurrentFolder))
                {
                    Logger.Debug("Processed source folder " + importState.CurrentFolder + " [=> " + importState.CurrentDestinationFolder +
                                 "], read=" + importState.CurrentFolderConsumed + ", imported=" + importState.CurrentFolderProcessed);
                }
                importState.CurrentFolder = msg.SourceFolder;
                importState.CurrentDestinationFolder = msg.DestinationFolder;
                importState.CurrentFolderProcessed = 0;
                importState.CurrentFolderConsumed = 0;
            }
            importState.CurrentFolderConsumed++;
            if (!importState.Stats.ContainsKey(msg.DestinationFolder))
            {
                importState.Stats.Add(msg.DestinationFolder, new MessageStats() { DestinationFolder = msg.DestinationFolder });
            }
            if (!importState.Stats[msg.DestinationFolder].SourceFolders.Contains(msg.SourceFolder))
            {
                importState.Stats[msg.DestinationFolder].SourceFolders.Add(msg.SourceFolder);
            }
            bool noDelete = false;
            try
            {
                ImportIntoOutlook(msg, _rootFolder.FolderPath + @"\" + msg.DestinationFolder);
                importState.CurrentFolderProcessed++;
                importState.Stats[msg.DestinationFolder].Count++;
                SucceededMessageCount++;
                //RecordImport(msg);
            }
            catch (Exception ex)
            {
                Logger.Error("Failed to import msg [" + msg + "] into outlook : " + ex.Message, ex);
                FailedMessageCount++;
                noDelete = true;
            }
            finally
            {
                ProcessedMessageCount++;
                if (!noDelete)
                {
                    try
                    {
                        File.Delete(msg.MsgFile);
                    }
                    catch (Exception ex)
                    {
                        Logger.Warn("Failed to delete msg file [" + msg.MsgFile + "] for uid [" + msg.SourceId + "] : " +
                                    ex.Message);
                    }
                }
            }
            _lastState = importState;
            return importState;
        }

        private void RecordImport(MsgDescriptor msg)
        {
            if (!String.IsNullOrEmpty(msg.SourceId))
            {
                try
                {
                    ImportIdsMsg.Body += msg.SourceId + "\n";
                    ImportIdsMsg.Save();
                }
                catch (Exception e)
                {
                    Logger.Error("Failed to record successful import of " + msg.SourceId + " : " + e.Message);
                }
            }
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

        #endregion

        #region Outlook Methods

        private Store PstHandle(NameSpace ns, string pstPath, string pstName)
        {
            ns.AddStore(pstPath);
            foreach (Store store in ns.Stores)
            {
                if (store.FilePath != null && pstPath.ToLower().Equals(store.FilePath.ToLower()))
                {
                    store.GetRootFolder().Name = pstName;
                    return store;
                }
            }
            throw new Exception("Failed to find store " + pstPath);
        }

        private void ImportIntoOutlook(MsgDescriptor msg, string folder)
        {
            bool failOnce = false;
            dynamic item = null;
            while (true)
            {
                // we always attempt this twice, sometimes it fails first time
                try
                {
                    // This may be a mailitem, we don't know, so don't convert
                    item = _outlook.GetNamespace("MAPI").OpenSharedItem(msg.MsgFile);
                    break;
                }
                catch (Exception e)
                {
                    if (failOnce) throw;
                    Logger.Warn("Opening Msg file failed [" + msg.MsgFile + "], retrying : " + e.Message);
                    failOnce = true;
                }
            }
            // Here we load the various meta data
            try
            {
                // We can't process this message if it has been encrypted/signed
                if (msg.IsEncrypted)
                {
                    throw new Exception("Cannot save message " + folder + @"\" + item.Subject + ", message is encrypted.");
                }
                else
                {
                    Logger.Debug("Importing " + folder + @"\" + item.Subject + " (" + String.Join(", ", msg.Flags) + ") [" + String.Join(", ", msg.Categories) + "]");

                    // This is required to be set one way or another, otherwise the message is marked as new (not delivered)
                    item.UnRead = msg.Flags.Contains(MessageFlags.Unread);

                    foreach (var messageFlag in msg.Flags)
                    {
                        try
                        {
                            switch (messageFlag)
                            {
                                case MessageFlags.FollowUp:
                                    {
                                        item.FlagRequest = "Follow Up";
                                        break;
                                    }
                                case MessageFlags.ReminderSet:
                                    {
                                        item.ReminderSet = true;
                                        break;
                                    }
                                case MessageFlags.ReadReceiptRequested:
                                    {
                                        item.ReadReceiptRequested = true;
                                        break;
                                    }
                                case MessageFlags.DeliveryReceiptRequested:
                                    {
                                        item.OriginatorDeliveryReportRequested = true;
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
                        if (msg.Importance != null) item.Importance = (OlImportance)msg.Importance;
                        if (msg.Sensitivity != null) item.Sensitivity = (OlSensitivity)msg.Sensitivity;
                        //item.TaskDueDate = msg.ReminderDueBy;
                        item.Categories = String.Join(",", msg.Categories);
                    }
                    catch (Exception e)
                    {
                        Logger.Warn(
                            "Failed to set metadata on " + folder + @"\" + item.Subject + ", ignoring metadata.", e);
                    }
                    item.Move(GetCreateFolder(folder, _outlook));
                }
            }
            finally
            {
                try
                {
                    item.Close(OlInspectorClose.olDiscard);
                }
                catch (Exception e)
                {
                    Logger.Warn("Failed to close " + item.Subject + "[" + msg.SourceId + "]");
                }

            }
        }

        private Folder GetCreateFolder(string folderPath, Application app)
        {
            Folder folder;
            string backslash = @"\";

            if (folderPath.StartsWith(@"\\"))
            {
                folderPath = folderPath.Remove(0, 2);
            }
            try
            {
                String[] folders = folderPath.Split(backslash.ToCharArray());
                try
                {
                    folder = app.Session.Folders[folders[0]] as Folder;
                }
                catch (Exception e)
                {
                    Logger.Error("Failed to select folder " + folders[0] + ", " + e.Message, e);
                    throw;
                }
                if (folder != null)
                {
                    for (int i = 1; i <= folders.GetUpperBound(0); i++)
                    {
                        Folders subFolders = folder.Folders;
                        try
                        {
                            folder = subFolders[folders[i]] as Folder;
                        }
                        catch
                        {
                            folder = null;
                        }
                        if (folder == null)
                        {
                            Logger.Debug("Creating folder " + folders[i] + " for " + "'" + folderPath + "'");
                            subFolders.Add(folders[i], Type.Missing);
                            Thread.Sleep(2000);
                            folder = subFolders[folders[i]] as Folder;
                            Logger.Debug("Created folder " + folders[i] + " for " + "'" + folderPath + "'");
                        }
                    }
                }
            }
            catch (Exception e)
            {
                Logger.Error("Failed to get/create folder for " + folderPath + " : " + e.Message);
                throw;
            }
            return folder;
        }

        #endregion

        public override void Close()
        {
            if (!Closed && Status != MessageProcessorStatus.Idle)
            {
                Closed = true;
                if (_queue != null) { _queue.Close(); }
                try
                {
                    _outlook.GetNamespace("MAPI").RemoveStore(_pstHandle.GetRootFolder());
                }
                catch
                {
                    Logger.Error("Failed to unmount pst");
                }
                Logger.Debug("Processed source folder " + _lastState.CurrentFolder + " [=> " + _lastState.CurrentDestinationFolder + "], read=" + _lastState.CurrentFolderConsumed + ", imported=" + _lastState.CurrentFolderProcessed);
                foreach (var messageStats in _lastState.Stats)
                {
                    Logger.Info(messageStats.Value.ToString());
                }
            }
        }
    }
}
