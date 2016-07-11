using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using Microsoft.Exchange.WebServices.Data;

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

    public enum FlagIcon
    {
        None,
        Outlook2003Purple,
        Outlook2003Orange,
        Outlook2003Green,
        Outlook2003Yellow,
        Outlook2003Blue,
        Outlook2003Red,
    }

    public class FollowUpFlag
    {
        public DateTime CompleteDateTime;
        public DateTime DueDateTime;
        public DateTime StartDateTime;
        public FollowUpFlagStatus Status;
    }

    public enum FollowUpFlagStatus
    {
        NotFlagged,
        Flagged,
        Complete,
    }

    public class BaseMessageDescriptor
    {
        public String SourceId;
        public DateTime? ReceivedDateTime;
        public DateTime? SentDateTime;
        public String SourceFolder;
        public String DestinationFolder;
        public FollowUpFlag FollowUpFlag; 
        public Collection<MessageFlags> Flags = new Collection<MessageFlags>();
        public FlagIcon FlagIcon { get; set; }
        public string Subject { get; set; }
        public string ItemClass { get; set; }
        public bool IsPublicFolder { get; set; }

        public Collection<string> Categories = new Collection<string>();
        public String ConversationId { get; set; }
        public int? IconIndex { get; set; }
        public int? Importance { get; set; }
        public string InReplyTo { get; set; }
        public int? Sensitivity { get; set; }
        public DateTime? ReminderDueBy { get; set; }
        public bool IsEncrypted { get; set; }

        public BaseMessageDescriptor()
        {
            FlagIcon = FlagIcon.None;
        }

        public override string ToString()
        {
            return "id=" + SourceId + ", sourceFolder=" + SourceFolder + ", destinationFolder=" + DestinationFolder;
        }
    }
}