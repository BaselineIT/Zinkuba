using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Zinkuba.MailModule;
using Zinkuba.MailModule.API;

namespace Zinkuba.App.Mailbox
{
    public interface IMailbox
    {
        IMessageSource GetSource();
        void StartExporter(MessagePipeline exporter);
        String Id { get; set; }
    }
}
