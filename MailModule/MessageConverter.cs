using System.Linq;
using S22.Imap;
using Zinkuba.MailModule.MessageDescriptor;

namespace Zinkuba.MailModule
{
    public class MessageConverter
    {
        public static MsgDescriptor ToMsgDescriptor(RawMessageDescriptor message)
        {
            MsgDescriptor msg = new MsgDescriptor();
            PopulateMessageDescriptior(message,msg);
            return msg;
        }

        private static void PopulateMessageDescriptior(BaseMessageDescriptor source, BaseMessageDescriptor target)
        {
            target.DestinationFolder = source.DestinationFolder;
            target.SourceFolder = source.SourceFolder;
            target.SourceId = source.SourceId;
            target.ConversationId = source.ConversationId;
            target.IconIndex = source.IconIndex;
            target.Importance = source.Importance;
            target.InReplyTo = source.InReplyTo;
            target.ReminderDueBy = source.ReminderDueBy;
            target.Sensitivity = source.Sensitivity;
            target.IsEncrypted = source.IsEncrypted;
            foreach (var category in source.Categories)
            {
                target.Categories.Add(category);
            }
            foreach (var messageFlag in source.Flags)
            {
                target.Flags.Add(messageFlag);
            }
        }
    }
}