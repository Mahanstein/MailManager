using System.Windows.Forms;

namespace MailManager
{
    public partial class frmEmails : Form
    {
        public frmEmails()
        {
            InitializeComponent();
            init();
        }

        private void init()
        {
            MailService.Service oMailService = new MailService.Service();
            oMailService.MonitoredMailboxes.Add(new MailService.MonitoredMailbox("rafas.rafaelsilva@hotmail.com", "TesteManagerIn", "TesteManagerProcessed"));
        }
    }
}
