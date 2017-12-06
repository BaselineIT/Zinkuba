using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using log4net;

namespace Zinkuba.MailModule
{
    public enum MailProvider
    {
        GmailImap,
        DefaultImap,
        Unknown,
        Exchange
    }


    class FolderMapping
    {
        private static readonly ILog Logger = LogManager.GetLogger(typeof(FolderMapping));
        public String Source;
        public String Destination;
        public bool Ignore = false;

        private static readonly Dictionary<MailProvider, List<FolderMapping>> ProviderMappings = new Dictionary
            <MailProvider, List<FolderMapping>>()
        {
            {
                MailProvider.GmailImap, new List<FolderMapping>
                {
                    new FolderMapping() { Source = @"[Gmail]", Ignore = true},
                    new FolderMapping() { Source = @"[Gmail]/Important", Ignore = true},
                    new FolderMapping() { Source = @"[Gmail]/Starred", Ignore = true},
                    new FolderMapping() { Source = @"[Gmail]/All Mail", Ignore = true},
                    new FolderMapping() { Source = @"[Gmail]/Sent Mail", Destination = "Sent Items"},
                    new FolderMapping() { Source = @"[Gmail]/Spam", Destination = "Junk E-mail"},
                    new FolderMapping() { Source = @"[Gmail]/Drafts", Destination = "Drafts"},
                    new FolderMapping() { Source = @"[Gmail]/Trash", Destination = "Deleted Items"},
                    new FolderMapping() { Source = @"[Gmail]/Bin", Destination = "Deleted Items"},
                    new FolderMapping() { Source = @"[Google Mail]", Ignore = true},
                    new FolderMapping() { Source = @"[Google Mail]/Important", Ignore = true},
                    new FolderMapping() { Source = @"[Google Mail]/Starred", Ignore = true},
                    new FolderMapping() { Source = @"[Google Mail]/All Mail", Ignore = true},
                    new FolderMapping() { Source = @"[Google Mail]/Sent Mail", Destination = "Sent Items"},
                    new FolderMapping() { Source = @"[Google Mail]/Spam", Destination = "Junk E-mail"},
                    new FolderMapping() { Source = @"[Google Mail]/Drafts", Destination = "Drafts"},
                    new FolderMapping() { Source = @"[Google Mail]/Trash", Destination = "Deleted Items"},
                    new FolderMapping() { Source = @"[Google Mail]/Bin", Destination = "Deleted Items"},
                    new FolderMapping() { SourceRegex = "^INBOX", DestinationRegex = "Inbox"},
                    new FolderMapping() { SourceRegex = "/", DestinationRegex = @"\" },
                }
            },
            {
                MailProvider.Unknown, new List<FolderMapping>()
            },
            {
                MailProvider.DefaultImap, new List<FolderMapping>
                {
                    new FolderMapping() { Source = @"Sent", Destination = "Sent Items"},
                    new FolderMapping() { Source = @"Sent Messages", Destination = "Sent Items"},
                    new FolderMapping() { Source = @"Deleted Messages", Destination = "Deleted Items"},
                    new FolderMapping() { Source = @"Trash", Destination = "Deleted Items"},
                    new FolderMapping() { SourceRegex = "/", DestinationRegex = @"\" },
                    new FolderMapping() { SourceRegex = "^INBOX", DestinationRegex = "Inbox"},
                }                
            },
            {
                MailProvider.Exchange, new List<FolderMapping>
                {
                    new FolderMapping() { SourceRegex = @"^Calendar($|\\*)", Ignore = true },
                    new FolderMapping() { SourceRegex = @"^Tasks($|\\*)", Ignore = true },
                    new FolderMapping() { Source = @"Conversation Action Settings", Ignore = true },
                    new FolderMapping() { SourceRegex = @"^Contacts($|\\*)", Ignore = true },
                    new FolderMapping() { Source = @"Suggested Contacts", Ignore = true },
                    new FolderMapping() { Source = @"ExternalContacts", Ignore = true },
                    new FolderMapping() { Source = @"Files", Ignore = true },
                    new FolderMapping() { Source = @"Outbox", Ignore = true },
                    new FolderMapping() { Source = @"PersonMetadata", Ignore = true },
                    new FolderMapping() { Source = @"Quick Step Settings", Ignore = true },
                    new FolderMapping() { Source = @"RSS Feeds", Ignore = true },
                    new FolderMapping() { SourceRegex = @"^Sync Issues($|\\*)", Ignore = true },
                }
            }
        };

        public string DestinationRegex { get; set; }

        public string SourceRegex { get; set; }

        public static List<FolderMapping> GetMappings(MailProvider provider)
        {
            return ProviderMappings[provider];
        }

        public static String ApplyMappings(String folder, List<FolderMapping> mappings)
        {
            if (String.IsNullOrWhiteSpace(folder)) { return null; }
            String previousFolder = folder;
            String returnFolder = folder;
            foreach (var mapping in mappings)
            {
                if ((String.IsNullOrEmpty(mapping.Source) || !mapping.Source.Equals(returnFolder)) &&
                    (String.IsNullOrEmpty(mapping.SourceRegex) || !Regex.Match(returnFolder, mapping.SourceRegex).Success)) continue;

                try
                {
                    previousFolder = returnFolder;
                    if (mapping.Ignore)
                    {
                        Logger.Debug("Ignoring folder " + returnFolder + (!String.IsNullOrEmpty(mapping.SourceRegex) ? "(matches " + mapping.SourceRegex + ")" : ""));
                        returnFolder = null;
                        break;
                    }
                    else
                    {
                        if (!String.IsNullOrEmpty(mapping.Source) && mapping.Source.Equals(returnFolder))
                        {
                            returnFolder = mapping.Destination;
                            Logger.Debug("Mapped " + previousFolder + " => " + returnFolder + " [" + folder + "]");
                        }
                        else
                        {
                            returnFolder = Regex.Replace(returnFolder, mapping.SourceRegex, mapping.DestinationRegex);
                            Logger.Debug("Mapped " + previousFolder + " => " + returnFolder + " [" + folder + "]");
                        }
                    }
                }
                catch (Exception ex)
                {
                    Logger.Error("Failed to apply mapping to " + folder, ex);
                }
            }
            return returnFolder;
        }

        public static String ApplyMappings(String folder, MailProvider provider)
        {
            return ApplyMappings(folder, GetMappings(provider));
        }
    }
}
