using System.Windows.Forms;

namespace MailManager
{
    public partial class frmAlreadyExecuting : Form
    {
        public frmAlreadyExecuting()
        {
            InitializeComponent();
        }

        private void frmAlreadyExecuting_FormClosed(object sender, FormClosedEventArgs e)
        {
            Application.Exit();
        }
    }
}
