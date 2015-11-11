using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using log4net;
using Microsoft.Exchange.WebServices.Data;
using Zinkuba.MailModule.API;

namespace Zinkuba.MailModule.MessageProcessor
{
    internal class ExchangeHelper
    {
        private static readonly ILog Logger = LogManager.GetLogger(typeof(ExchangeHelper));

        internal static readonly ExchangeVersion[] ExchangeVersions = { ExchangeVersion.Exchange2013, ExchangeVersion.Exchange2010_SP2, ExchangeVersion.Exchange2010_SP1, ExchangeVersion.Exchange2010, ExchangeVersion.Exchange2007_SP1 };

        internal static ExchangeService ExchangeConnect(String hostname, String username, String password)
        {
            ServicePointManager.ServerCertificateValidationCallback = CertificateValidationCallBack;
            int attempt = 0;
            ExchangeService exchangeService = null;
            do
            {
                try
                {
                    exchangeService = new ExchangeService(ExchangeHelper.ExchangeVersions[attempt])
                    {
                        Credentials = new WebCredentials(username, password),
                        Url = new Uri("https://" + hostname + "/EWS/Exchange.asmx"),
                        Timeout = 30*60*1000, // 30 mins
                    };
                    Logger.Debug("Binding to exchange server " + exchangeService.Url + " as " + username + ", version " +
                                 ExchangeHelper.ExchangeVersions[attempt]);
                    Folder.Bind(exchangeService, WellKnownFolderName.MsgFolderRoot);
                }
                catch (ServiceVersionException e)
                {
                    Logger.Warn("Failed to bind as version " + ExchangeHelper.ExchangeVersions[attempt]);
                    exchangeService = null;
                    attempt++;
                }
                catch (Exception e)
                {
                    Logger.Error("Failed to bind to exchange server", e);
                    if (e.Message.Contains("Unauthorized"))
                    {
                        throw new MessageProcessorException(e.Message) { Status = MessageProcessorStatus.AuthFailure };
                    }
                    throw new MessageProcessorException(e.Message) { Status = MessageProcessorStatus.ConnectionError };
                }
            } while (exchangeService == null && attempt < ExchangeHelper.ExchangeVersions.Count());
            if (exchangeService == null)
            {
                throw new MessageProcessorException("Failed to connect to " + hostname + " with username " + username)
                {
                    Status = MessageProcessorStatus.ConnectionError
                };
            }
            return exchangeService;
        }

        private static bool CertificateValidationCallBack(
            object sender,
            System.Security.Cryptography.X509Certificates.X509Certificate certificate,
            System.Security.Cryptography.X509Certificates.X509Chain chain,
            System.Net.Security.SslPolicyErrors sslPolicyErrors)
        {
            // If the certificate is a valid, signed certificate, return true.
            if (sslPolicyErrors == System.Net.Security.SslPolicyErrors.None)
            {
                return true;
            }

            // If there are errors in the certificate chain, look at each error to determine the cause.
            if ((sslPolicyErrors & System.Net.Security.SslPolicyErrors.RemoteCertificateChainErrors) != 0)
            {
                if (chain != null && chain.ChainStatus != null)
                {
                    foreach (System.Security.Cryptography.X509Certificates.X509ChainStatus status in chain.ChainStatus)
                    {
                        if ((certificate.Subject == certificate.Issuer) &&
                            (status.Status == System.Security.Cryptography.X509Certificates.X509ChainStatusFlags.UntrustedRoot))
                        {
                            // Self-signed certificates with an untrusted root are valid. 
                            //Logger.Warn("Self signed certificate, continuing regardless.");
                            continue;
                        }
                        else if (status.Status == System.Security.Cryptography.X509Certificates.X509ChainStatusFlags.NotTimeValid)
                        {
                            // expired, we don't mind
                            //Logger.Warn("Certificate has expired, continuing regardless.");
                            continue;
                        }
                        else if (status.Status == System.Security.Cryptography.X509Certificates.X509ChainStatusFlags.PartialChain)
                        {
                            // chain has an invalid or inaccessible root cert, we don't mind this either (badly configured local exchanges)
                            //Logger.Warn("Certificate chain is partial, continuing regardless.");
                            continue;
                        }
                        else
                        {
                            if (status.Status !=
                                System.Security.Cryptography.X509Certificates.X509ChainStatusFlags.NoError)
                            {
                                // If there are any other errors in the certificate chain, the certificate is invalid,
                                // so the method returns false.
                                return false;
                            }
                        }
                    }
                }

                // When processing reaches this line, the only errors in the certificate chain are 
                // untrusted root errors for self-signed certificates. These certificates are valid
                // for default Exchange server installations, so return true.
                return true;
            } else if ((sslPolicyErrors & System.Net.Security.SslPolicyErrors.RemoteCertificateNameMismatch) != 0)
            {
               // Certificate name is not correct, we don't care
                return true;
            }
            else
            {
                // In all other cases, return false.
                return false;
            }
        }


        public static void GetFolderSummary(ExchangeService service, List<ExchangeFolder> folderStore, DateTime startDate, DateTime endDate, bool purgeIgnored = true)
        {
            SearchFilter.SearchFilterCollection filter = new SearchFilter.SearchFilterCollection();
            filter.LogicalOperator = LogicalOperator.And;
            Logger.Debug("Getting mails from " + startDate + " to " + endDate);
            filter.Add(new SearchFilter.IsGreaterThanOrEqualTo(ItemSchema.DateTimeReceived, startDate));
            filter.Add(new SearchFilter.IsLessThanOrEqualTo(ItemSchema.DateTimeReceived, endDate));
            var view = new ItemView(20, 0, OffsetBasePoint.Beginning) { PropertySet = PropertySet.IdOnly };
            var ignoredFolders = new List<ExchangeFolder>();
            foreach (var exchangeFolder in folderStore)
            {
                var destinationFolder = FolderMapping.ApplyMappings(exchangeFolder.FolderPath, MailProvider.Exchange);
                if (!String.IsNullOrWhiteSpace(destinationFolder))
                {
                    exchangeFolder.MappedDestination = destinationFolder;
                    var findResults = service.FindItems(exchangeFolder.FolderId, filter, view);
                    Logger.Debug(exchangeFolder.FolderPath + " => " + exchangeFolder.MappedDestination + ", " +
                                 findResults.TotalCount + " messages.");
                    exchangeFolder.MessageCount = findResults.TotalCount;
                }
                else
                {
                    ignoredFolders.Add(exchangeFolder);
                }
            }
            if (purgeIgnored)
            {
                foreach (var exchangeFolder in ignoredFolders)
                {
                    folderStore.Remove(exchangeFolder);
                }
            }
        }

        public static void GetAllFolders(ExchangeService service, ExchangeFolder currentFolder, List<ExchangeFolder> folderStore, bool skipEmpty = true)
        {
            Logger.Debug("Looking for sub folders of '" + currentFolder.FolderPath + "'");
            var results = service.FindFolders(currentFolder.Folder.Id, new FolderView(int.MaxValue));
            foreach (var folder in results)
            {
                String folderPath = (String.IsNullOrEmpty(currentFolder.FolderPath)
                    ? ""
                    : currentFolder.FolderPath + @"\") + folder.DisplayName;
                if (skipEmpty && folder.TotalCount == 0 && folder.ChildFolderCount == 0)
                {
                    Logger.Debug("Skipping folder " + folderPath + ", no messages, no subfolders.");
                    continue;
                }
                Logger.Debug("Found folder " + folderPath + ", " + folder.TotalCount + " messages in total.");
                var exchangeFolder = new ExchangeFolder()
                {
                    Folder = folder,
                    FolderPath = folderPath,
                    MessageCount = folder.TotalCount,
                    FolderId = folder.Id,
                };
                folderStore.Add(exchangeFolder);
                if (exchangeFolder.Folder.ChildFolderCount > 0)
                {
                    GetAllFolders(service, exchangeFolder, folderStore, skipEmpty);
                }
            }
        }

    }


}