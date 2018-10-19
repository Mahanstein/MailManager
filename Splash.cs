using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;

namespace MailManager
{
    public partial class Splash : Form
    {
        public Splash()
        {
            InitializeComponent();
        }

        private void VerifyExecution()
        {
            string appGuid = ((GuidAttribute)Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(GuidAttribute), false).GetValue(0)).Value.ToString();

            string mutexId = string.Format("Global\\{{{0}}}", appGuid);

            Mutex mut = new Mutex(false, mutexId);
            if(mut.WaitOne(3000))
            {
                start();
            }
            else
            {
                frmAlreadyExecuting fExecuting = new frmAlreadyExecuting();
                Hide();
                fExecuting.Show();
            }

        }

        private void start()
        {
            Logging.Logger.SetPath(@"C:\LogTest\");
            Logging.Logger.Log("Iniciando.");

            Hide();
            Program.ntIcon.Visible = true;

            //frmMonitor fMonior = new frmMonitor();
            //fMonior.Show();
        }

        private void Splash_Shown(object sender, System.EventArgs e)
        {
            Application.DoEvents();
            VerifyExecution();
        }
    }
}
