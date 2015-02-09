using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace Zinkuba.MailModule.MessageDescriptor
{

    public enum MessageFlags
    {
        Unread,
        FollowUp,
        Draft,
        ReadReceiptRequested,
        DeliveryReceiptRequested,
        ReminderSet,
        Associated,
        FromMe,
        Resend
    }

    public class BaseMessageDescriptor
    {
        public String SourceId;
        public String SourceFolder;
        public String DestinationFolder;
        public Collection<MessageFlags> Flags = new Collection<MessageFlags>();
        public string Subject { get; set; }

        public Collection<string> Categories = new Collection<string>();
        public String ConversationId { get; set; }
        public int? IconIndex { get; set; }
        public int? Importance { get; set; }
        public string InReplyTo { get; set; }
        public int? Sensitivity { get; set; }
        public DateTime? ReminderDueBy { get; set; }
        public bool IsEncrypted { get; set; }

        public override string ToString()
        {
            return "id=" + SourceId + ", sourceFolder=" + SourceFolder + ", destinationFolder=" + DestinationFolder;
        }
    }
}