using System;
using System.Collections.Generic;

namespace Zinkuba.MailModule
{
    public class MessageStats
    {
        internal List<string> SourceFolders = new List<string>();
        internal String DestinationFolder;
        internal int Count;

        public override string ToString()
        {
            return "DestinationFolder=" + DestinationFolder + ", SourceFolders=[" + String.Join(",", SourceFolders) +
                   "], Count=" + Count;

        }
    }
}