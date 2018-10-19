using Microsoft.Exchange.WebServices.Data;

namespace MailManager.MailService
{
    class MonitoredMailbox
    {
        public string Mailbox { get; set; }

        public string OrginalFolderName { get; set; }

        public Folder OriginalFolder { get; set; }

        public string DestinationFolderName { get; set; }

        public Folder DestinationFolder { get; set; }

        public MonitoredMailbox(string pMailbox, string pOrigin, string pDestination)
        {
            Mailbox = pMailbox;
            OrginalFolderName = pOrigin;
            DestinationFolderName = pDestination;
        }
    }
}
