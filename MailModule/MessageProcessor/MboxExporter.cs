using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using log4net;
using Zinkuba.MailModule.API;
using Zinkuba.MailModule.MessageDescriptor;
using Exception = System.Exception;

namespace Zinkuba.MailModule.MessageProcessor
{
    public class MboxExporter : BaseMessageProcessor, IMessageWriter<RawMessageDescriptor>, IMessageSource
    {
        private readonly Dictionary<string, string> _folderMap;
        private IMessageReader<RawMessageDescriptor> _nextReader;
        private static readonly ILog Logger = LogManager.GetLogger(typeof(MboxExporter));

        public IMessageReader<RawMessageDescriptor> NextReader
        {
            set { _nextReader = value; }
            get
            {
                return _nextReader;
            }
        }

        public bool TestOnly { get; set; }

        private Thread _sourceThread;

        public int TotalMessages
        {
            get { return _totalMessages; }
            private set { _totalMessages = value; OnTotalMesssagesChanged(); }
        }

        private int _totalMessages;
        private int _nextUid;
        private StringBuilder _messageBuffer;

        public MboxExporter(String name, Dictionary<String,String> folderMap)
        {
            _folderMap = folderMap;
            Name = name;
        }

        public override void Initialise()
        {
            Status = MessageProcessorStatus.Initialising;
            /*
            try
            {
                _filestream = new StreamReader(_mboxPath);
            }
            catch (Exception ex)
            {
                Logger.Error("Opening mbox file [" + _mboxPath + "] failed : " + ex.Message, ex);
                Status = MessageProcessorStatus.UnknownError;
                throw ex;
            }
             */
            Status = MessageProcessorStatus.Initialised;
        }

        public void Start()
        {
            Initialise();
            _sourceThread = new Thread(RunMboxReader) { IsBackground = true, Name = "sourceThread-" + Name };
            _sourceThread.Start();
        }

        public event EventHandler TotalMessagesChanged;

        protected virtual void OnTotalMesssagesChanged()
        {
            EventHandler handler = TotalMessagesChanged;
            if (handler != null) handler(this, EventArgs.Empty);
        }

        private void RunMboxReader()
        {
            Status = MessageProcessorStatus.Started;
            _messageBuffer = new StringBuilder();
            char[] buffer = new char[1024 * 1024];
            foreach (var folderMap in _folderMap)
            {
                StreamReader filestream = null;
                String mboxPath = folderMap.Key;
                String folderPath = folderMap.Value;
                try
                {
                    filestream = new StreamReader(mboxPath);
                    // drop the first From line.
                    filestream.ReadLine();
                    int read = filestream.Read(buffer, 0, buffer.Length);
                    while (read > 0)
                    {
                        _messageBuffer.Append(buffer, 0, read);
                        //Logger.Debug(buffer);
                        var eml = GetNextMessage(filestream);
                        while (!String.IsNullOrEmpty(eml))
                        {
                            try
                            {
                                var messageDescriptor = buildMessage(eml, folderPath);
                                _nextReader.Process(messageDescriptor);
                                SucceededMessageCount++;
                            }
                            catch (Exception ex)
                            {
                                Logger.Error("Failed to get and enqueue Imap message [" + _nextUid + "]", ex);
                                FailedMessageCount++;
                            }
                            ProcessedMessageCount++;
                            eml = GetNextMessage(filestream);
                        }
                        read = filestream.Read(buffer, 0, buffer.Length);
                        if (TestOnly && ProcessedMessageCount > 20)
                        {
                            break;
                        }
                    }
                }
                catch (Exception ex)
                {
                    Logger.Error("Opening mbox file [" + mboxPath + "] failed : " + ex.Message, ex);
                    throw ex;
                }
                finally
                {
                    Logger.Info("Completed folder " + folderPath + ", " + ProcessedMessageCount + " processed, " +
                                SucceededMessageCount + " succeeded, " + IgnoredMessageCount + " ignored, " +
                                FailedMessageCount + " failed.");
                    try
                    {
                        if (filestream != null) filestream.Close();
                    }
                    catch (Exception e)
                    {
                        Logger.Warn("Failed to close file " + mboxPath, e);
                    }
                }
            }
            Close();
        }

        private string GetNextMessage(StreamReader filestream)
        {
            Match match = Regex.Match(_messageBuffer.ToString(), @"[\r\n]+From .*?[\r\n]+");
            String eml = null;
            if (match.Success)
            {
                //Logger.Debug("Found " + match.ToString());
                eml = _messageBuffer.ToString(0, match.Index);
                _messageBuffer.Remove(0, match.Index + match.Length);
                //Logger.Debug("Got message, more to come");
            }
            else if (filestream.Peek() < 0)
            {
                eml = _messageBuffer.ToString();
                _messageBuffer.Remove(0, _messageBuffer.Length);
                //Logger.Debug("Got message, last one");
            }
            return eml;
        }

        private RawMessageDescriptor buildMessage(String eml, String folder)
        {
            var messageDescriptor = new RawMessageDescriptor()
            {
                RawMessage = Regex.Replace(Regex.Replace(eml, @"\n>(>*From)", "\n$1"), @"\r{0,1}\n", "\r\n"),
                DestinationFolder = folder,
                SourceFolder = folder,
                SourceId = "" + _nextUid++
            };
            return messageDescriptor;
        }

        public override void Close()
        {
            if (!Closed)
            {
                Closed = true;
                _nextReader.Close();
            }
        }

        public Type OutMessageDescriptorType()
        {
            return typeof (RawMessageDescriptor);
        }
    }
}
