using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Zinkuba.MailModule
{
    public class MailFolder
    {
        public String SourceFolder { get; set; }
        public String DestinationFolder { get; set; }
        public int MessageCount { get; set; }
    }
}
