using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;

namespace MailManager.MailService
{
    class Service
    {
        private ExchangeService ExService;

        public List<MonitoredMailbox> MonitoredMailboxes = new List<MonitoredMailbox>();

        public bool Connect(string pUser, string pPassword)
        {
            try
            {
                Logging.Logger.Log("Conectando no exchange.");

                ExService = new ExchangeService();

                NetworkCredential credential = new NetworkCredential(pUser, pPassword);
                ExService.UseDefaultCredentials = false;
                ExService.Credentials = credential;
                
                ExService.AutodiscoverUrl(pUser);

                if (ExService.Url == null || string.IsNullOrEmpty(ExService.Url.AbsolutePath))
                    throw new Exception(Properties.Resources.ErrExchangeAutoDiscoverFail);

                CheckFolders();
                return true;
            }
            catch(Exception ex)
            {
                string err = Properties.Resources.ErrExchangeConnection;
                Logging.Logger.Log(err);
                return false;
            }
        }

        private void CheckFolders()
        {
            try
            {
                foreach (MonitoredMailbox item in MonitoredMailboxes)
                {
                    Mailbox mb = new Mailbox(item.Mailbox);

                    FolderId rootId;
                    if (mb.IsValid)
                        rootId = new FolderId(WellKnownFolderName.MsgFolderRoot, mb);
                    else
                        rootId = new FolderId(WellKnownFolderName.MsgFolderRoot);

                    FolderView vw = new FolderView(100000);
                    vw.Traversal = FolderTraversal.Deep;

                    FindFoldersResults result = ExService.FindFolders(rootId, vw);
                    item.OriginalFolder = result.Folders.FirstOrDefault(x => x.DisplayName == item.OrginalFolderName);
                    item.DestinationFolder = result.Folders.FirstOrDefault(x => x.DisplayName == item.DestinationFolderName);

                    if (item.OriginalFolder == null)
                        throw new Exception(Properties.Resources.ErrExchangeFolderNotFound + item.OrginalFolderName);

                    if (item.DestinationFolder == null)
                        throw new Exception(Properties.Resources.ErrExchangeFolderNotFound + item.DestinationFolderName);
                }
            }
            catch (Exception ex)
            {
                Logging.Logger.Log("Falha geral obtendo diretórios do exchange: " + ex.Message);
            }
        }

        public int CountEmails(string pFolder)
        {
            try
            {
                var folder = MonitoredMailboxes.FirstOrDefault(x => x.OrginalFolderName == pFolder);

                if (folder == null)
                    return -1;
                else
                    return folder.OriginalFolder.TotalCount;
            }
            catch (Exception ex)
            {
                Logging.Logger.Log("Falha geral obtendo contagem de itens: " + ex.Message);
                return 0;
            }
        }

        public List<string> DownloadEmails(Folder pFolderFrom)
        {
            string baseFolder = @"C:\Testes\Emails\" + pFolderFrom.DisplayName;

            List<string> result = new List<string>();
            try
            {
                int pageSize = 100;
                int total = pFolderFrom.TotalCount;
                int pos = 0;

                do
                {
                    ItemView vw = new ItemView(pageSize);
                    vw.Offset = pos;
                    FindItemsResults<Item> items = pFolderFrom.FindItems(vw);
                    foreach (EmailMessage item in items)
                    {
                        item.Load(new PropertySet(EmailMessageSchema.InternetMessageId, ItemSchema.MimeContent, EmailMessageSchema.Subject, EmailMessageSchema.DateTimeReceived, EmailMessageSchema.Sender));
                        MimeContent mc = item.MimeContent;

                        string folder = baseFolder + string.Format(@"\{0:yyyy}\{0:MM}\{0:dd}\", item.DateTimeReceived);
                        string fileName = createNameFromSender(folder, item.Sender);
                        string path = folder + fileName + ".eml";

                        Directory.CreateDirectory(folder);

                        if(File.Exists(path))
                        {
                            result.Add(item.Subject + " já existe: " + path);
                            Logging.Logger.Log(item.Subject + " já existe: " + path);
                        }
                        else
                        {
                            FileStream fs = new FileStream(path, FileMode.CreateNew);
                            fs.Write(mc.Content, 0, mc.Content.Length);
                            fs.Close();

                            result.Add(item.Subject + " salvo como " + path);
                            Logging.Logger.Log(item.Subject + " salvo como " + path);
                        }
                    }   
                    pos += pageSize;
                }
                while (pos <= total);

                return result;
            }
            catch (Exception ex)
            {
                Logging.Logger.Log("Falha baixando email: " + ex.Message);
                throw;
            }
        }

        private string createNameFromSender(string folder, EmailAddress sender)
        {
            int num = 1;
            string nome = sender.Address.Replace(".", "~");
            string fileName = string.Format("{0}_{1}", nome, num);
            while (File.Exists(folder + fileName))
            {
                num++;
                fileName = string.Format("{0}_{1}", nome, num);
            }
            return fileName;
        }

        public List<string> CopyEmails(Folder pFolderFrom, Folder pFolderTo)
        {
            List<string> result = new List<string>();

            try
            {
                int pageSize = 100;
                int total = pFolderFrom.TotalCount;
                int pos = 0;

                do
                {
                    ItemView vw = new ItemView(pageSize);
                    vw.Offset = pos;
                    FindItemsResults<Item> items = pFolderFrom.FindItems(vw);
                    foreach (EmailMessage item in items)
                    {
                        ItemView vwTo = new ItemView(10);
                        SearchFilter.SearchFilterCollection filter = new SearchFilter.SearchFilterCollection(LogicalOperator.And);
                        filter.Add(new SearchFilter.IsEqualTo(EmailMessageSchema.Subject, item.Subject));
                        filter.Add(new SearchFilter.IsEqualTo(EmailMessageSchema.InternetMessageId, item.InternetMessageId));
                        filter.Add(new SearchFilter.IsEqualTo(EmailMessageSchema.Size, item.Size));
                        filter.Add(new SearchFilter.IsEqualTo(EmailMessageSchema.From, item.From));
                        filter.Add(new SearchFilter.IsEqualTo(EmailMessageSchema.DisplayTo, item.DisplayTo));

                        var teste = pFolderTo.FindItems(filter, vwTo);
                        if (teste.Count() == 0)
                        {
                            item.Copy(pFolderTo.Id);
                            result.Add(item.Subject + " adicionado.");
                            Logging.Logger.Log(item.Subject + " adicionado.");
                        }
                        else
                        {
                            result.Add(item.Subject + " já presente.");
                            Logging.Logger.Log(item.Subject + " já presente.");
                        }
                    }
                    pos += pageSize;
                }
                while (pos <= total);

                return result;
            }
            catch (Exception ex)
            {
                result.Add(string.Format("Erro copiando de {0} para {1}: {2}", pFolderFrom.DisplayName, pFolderTo.DisplayName, ex.Message));
                Logging.Logger.Log(string.Format("Erro copiando de {0} para {1}: {2}", pFolderFrom.DisplayName, pFolderTo.DisplayName, ex.Message));
                return result;
            }
        }

    }
}
