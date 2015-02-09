using System;

namespace Zinkuba.MailModule.MessageDescriptor
{
    public class MsgDescriptor : BaseMessageDescriptor
    {
        public String MsgFile;

        public override string ToString()
        {
            return base.ToString() + ", file=" + MsgFile;
        }
    }
}