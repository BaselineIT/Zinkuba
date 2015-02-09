namespace Zinkuba.MailModule
{
    public class MessageProcessState : IPCState
    {
        public bool StartedNextConsumer { get; set; }
        public string CurrentFolder { get; set; }
        public int CurrentFolderConsumed { get; set; }
        public string CurrentDestinationFolder { get; set; }
        public int CurrentFolderProcessed { get; set; }

        public MessageProcessState()
        {
            CurrentFolder = "";
            CurrentDestinationFolder = "";
            CurrentFolderConsumed = 0;
            CurrentFolderProcessed = 0;
            StartedNextConsumer = false;
        }
    }
}