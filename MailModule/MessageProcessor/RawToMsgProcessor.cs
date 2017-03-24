using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;
using log4net;
using Microsoft.Office.Interop.Outlook;
using Microsoft.Win32;
using Zinkuba.MailModule.API;
using Zinkuba.MailModule.MessageDescriptor;
using Exception = System.Exception;

namespace Zinkuba.MailModule.MessageProcessor
{
    public class RawToMsgProcessor : BaseMessageProcessor, IMessageConnector<RawMessageDescriptor, MsgDescriptor>, IDisposable
    {
        private IMessageReader<MsgDescriptor> _nextReader;
        private static readonly ILog Logger = LogManager.GetLogger(typeof(RawToMsgProcessor));
        private PCQueue<RawMessageDescriptor, MessageProcessState> _queue;
        private MessageProcessState _lastState;
        private String _pathToMrMapi = "mrmapi.exe";

        public RawToMsgProcessor()
        {
            // Determine bitness of outlook
            var outlook = new Application();
            var bit = GetOutlookBitness(outlook);
            // Find the EXE
            String mrmapiResourceName = Assembly.GetExecutingAssembly().FullName.Split(',').First() + ".mrmapi_" + bit + ".exe";
            using (var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(mrmapiResourceName))
            {
                if (stream != null)
                {
                    Byte[] assemblyData = new Byte[stream.Length];
                    stream.Read(assemblyData, 0, assemblyData.Length);
                    String tempFile = Path.GetTempPath() + Guid.NewGuid() + "-mrmapi.exe";
                    File.WriteAllBytes(tempFile, assemblyData);
                    _pathToMrMapi = tempFile;
                    Logger.Debug("Using " + tempFile + " as mrmapi.exe (" + mrmapiResourceName + ")");
                }
            }
        }

        private static string GetOutlookBitness(Application outlook)
        {
            var version = outlook.Version.Split(new[] {'.'})[0];
            String bit = "32";
            var bitnessVariable =
                Registry.GetValue(@"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\" + version + @".0\Outlook","Bitness",null);
            if (bitnessVariable != null)
            {
                if ("x64".Equals(bitnessVariable))
                {
                    bit = "64";
                }
                else
                {
                    bit = "32";
                }
                Logger.Info("Outlook seems to be " + bit + "bit");
            }
            else
            {
                // then we have no idea
                Logger.Info("Failed to detect outlook app bitness, defaulting to " + bit);
            }
            /* This is the old way
            var path = Registry.GetValue(@"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\" + version + @".0\Word\InstallRoot",
                "Path", null);
            if (path == null)
            {
                path =
                    Registry.GetValue(
                        @"HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\Office\" + version + @".0\Word\InstallRoot",
                        "Path", null);
                if (path == null)
                {
                    // then we have no idea
                    Logger.Info("Failed to detect outlook app bitness, defaulting to " + bit);
                }
                else
                {
                    Logger.Info("Outlook seems to be 32bit on 64bit arch");
                    bit = "32"; // this seems to be 32bit on 64arch
                }
            }
            else
            {
                bit = "AMD64".Equals(Environment.GetEnvironmentVariable("PROCESSOR_ARCHITECTURE")) ? "64" : "32";
                Logger.Info("Outlook seems to be " + bit + "bit on " + bit + "bit arch");
            }
            */
            return bit;
        }


        public IMessageReader<MsgDescriptor> NextReader
        {
            get { return _nextReader; }
            set { _nextReader = value; }
        }

        public Type OutMessageDescriptorType()
        {
            return typeof(MsgDescriptor);
        }

        public Type InMessageDescriptorType()
        {
            return typeof(RawMessageDescriptor);
        }


        public void Process(RawMessageDescriptor messageDescriptor)
        {
            if (Status == MessageProcessorStatus.Initialised)
                Status = MessageProcessorStatus.Started;
            if (Status == MessageProcessorStatus.Started)
            {
                // limit the size of the queue
                while (_queue.Count > 150)
                {
                    Thread.Sleep(1000);
                }
                _queue.Consume(messageDescriptor);
            }
            else
            {
                throw new Exception("Cannot process, " + Name + " is not started.");
            }
        }

        public override void Initialise(List<MailFolder> folderList)
        {
            Status = MessageProcessorStatus.Initialising;
            NextReader.Initialise(folderList);
            _lastState = new MessageProcessState();
            _queue = new PCQueue<RawMessageDescriptor, MessageProcessState>(Name)
            {
                ProduceMethod = ProcessMessage,
                InitialiseProducer = () => _lastState,
                ShutdownProducer = ShutdownQueue
            };
            _queue.Start();
            Status = MessageProcessorStatus.Initialised;
            Logger.Info("RawToMessageProcessor Initialised");
        }

        private void ShutdownQueue(MessageProcessState state, Exception ex)
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

        private MessageProcessState ProcessMessage(RawMessageDescriptor message, MessageProcessState state)
        {
            if (!message.SourceFolder.Equals(state.CurrentFolder))
            {
                if (!String.IsNullOrWhiteSpace(state.CurrentFolder))
                {
                    Logger.Debug("Processed folder " + state.CurrentFolder + ", read=" + state.CurrentFolderConsumed + ", queued=" + state.CurrentFolderProcessed);
                }
                state.CurrentFolder = message.SourceFolder;
                state.CurrentFolderProcessed = 0;
                state.CurrentFolderConsumed = 0;
            }
            state.CurrentFolderConsumed++;
            try
            {
                /* This is now initialised in the initialise method
                if (_nextReader.Status == MessageProcessorStatus.Idle)
                {
                    _nextReader.Initialise();
                }
                 */
                _nextReader.Process(ConvertRawToMsg(message));
                state.CurrentFolderProcessed++;
                SucceededMessageCount++;
            }
            catch (Exception ex)
            {
                Logger.Error("Failed to process raw [" + message.SourceId + "] into msg : " + ex.Message, ex);
                FailedMessageCount++;
            }
            finally
            {
                ProcessedMessageCount++;
            }

            return state;
        }

        private MsgDescriptor ConvertRawToMsg(RawMessageDescriptor message)
        {
            MsgDescriptor msgDescriptor = MessageConverter.ToMsgDescriptor(message);
            // First commit the message to an eml file.
            String fileName = Path.GetTempFileName();
            Process process = null;
            try
            {
                var emlFileStream = new StreamWriter(fileName);
                emlFileStream.Write(message.RawMessage);
                emlFileStream.Close();
                msgDescriptor.MsgFile = Path.GetTempPath() + Guid.NewGuid() + ".msg";
                process = new Process
                {
                    StartInfo =
                    {
                        FileName = _pathToMrMapi,
                        Arguments = "-Ma -i \"" + fileName + "\" -o \"" + msgDescriptor.MsgFile + "\" -Cc CCSF_SMTP",
                        CreateNoWindow = true,
                        RedirectStandardOutput = true,
                        UseShellExecute = false
                    }
                };
                process.Start();
                if (!process.WaitForExit(30000))
                {
                    try
                    {
                        process.Kill();
                    }
                    catch (Exception ex)
                    {
                        Logger.Warn("Failed to kill process " + process.ProcessName);
                    }
                    throw new TimeoutException("Process failed to exit [" + process.ProcessName + "]");
                }
            }
            catch (Exception e)
            {
                Logger.Error("MrMapi failed to process file " + fileName,e);
                throw;
            }
            finally
            {
                if (process != null)
                {
                    try
                    {
                        String output = process.StandardOutput.ReadToEnd();
                        if (process.ExitCode != 0 || output.Contains("Conversion returned an error") || output.Contains("not found"))
                        {
                            string error = "MrMapi failed to process file " + fileName + " with exit code " +
                                           process.ExitCode + "\n" + output;
                            throw new Exception(error);
                        }
                    }
                    finally
                    {
                        process.Close();
                    }
                }
                try
                {
                    File.Delete(fileName);
                }
                catch (Exception ex)
                {
                    Logger.Warn("Failed to delete eml file [" + fileName + "] for uid [" + message.SourceId + "]", ex);
                }
            }
            return msgDescriptor;
        }

        public override void Close()
        {
            if (!Closed && Status != MessageProcessorStatus.Idle)
            {
                Closed = true;
                if (_queue != null) { _queue.Close(); }
                _nextReader.Close();
            }
        }

        public void Dispose()
        {
            RemoveMrMapi();
        }

        ~RawToMsgProcessor()
        {
            RemoveMrMapi();
        }

        private void RemoveMrMapi()
        {
            if (File.Exists(_pathToMrMapi))
            {
                try
                {
                    //File.Delete(_pathToMrMapi);
                    Logger.Debug("Removed " + _pathToMrMapi);
                }
                catch (Exception e)
                {
                    Logger.Warn("Failed to remove mrmapi", e);
                }
            }
        }
    }
}
