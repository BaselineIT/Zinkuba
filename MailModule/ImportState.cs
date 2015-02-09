using System.Collections.Generic;

namespace Zinkuba.MailModule
{
    public class ImportState : MessageProcessState
    {
        public Dictionary<string, MessageStats> Stats { get; set; }

        public ImportState() : base()
        {
            Stats = new Dictionary<string, MessageStats>();
        }
    }
}