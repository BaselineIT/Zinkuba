using System;
using System.Collections.ObjectModel;
using System.Linq;
using log4net;
using Zinkuba.App.Folder;
using Zinkuba.App.MailAccount;
using Zinkuba.MailModule;
using Zinkuba.MailModule.API;
using Zinkuba.MailModule.MessageProcessor;

namespace Zinkuba.App.Mailbox
{
    public class MboxMailbox : IMailbox
    {
        private static readonly ILog Logger = LogManager.GetLogger(typeof (MboxMailbox));
        private readonly MboxAccount _account;
        public String Name;
        public MessagePipeline Exporter { get; private set; }
        public event EventHandler ExportStarted;

        public ObservableCollection<MboxFolder> Folders;

        protected virtual void OnExportStarted()
        {
            EventHandler handler = ExportStarted;
            if (handler != null) handler(this, EventArgs.Empty);
        }

        public MboxMailbox(MboxAccount account)
        {
            _account = account;
            Folders = new ObservableCollection<MboxFolder>();
        }

        public IMessageSource GetSource()
        {
            return new MboxSource(Name, Folders.ToDictionary(folder => folder.MboxPath, folder => folder.FolderPath));
        }

        public void StartExporter(MessagePipeline exporter)
        {
            if (Exporter == null || Exporter.Failed)
            {
                Exporter = exporter;
                OnExportStarted();
                try
                {
                    exporter.Start();
                }
                catch (Exception ex)
                {
                    Logger.Error("Failed to start exporter", ex);
                }
            }
        }

        public string Id { get; set; }
    }
}
