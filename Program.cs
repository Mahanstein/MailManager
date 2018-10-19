using System;
using System.Windows.Forms;

namespace MailManager
{
    static class Program
    {
        public static NotifyIcon ntIcon = new NotifyIcon();

        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.ApplicationExit += Application_ApplicationExit;

            setupIcon();

            Splash fSplash = new Splash();
            fSplash.Show();

            Application.Run();
        }

        private static void Application_ApplicationExit(object sender, EventArgs e)
        {
            ntIcon.Visible = false;
            ntIcon.Dispose();
            Logging.Logger.Log("Saindo.");
        }

        private static void sair_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private static void monitor_Click(object sender, EventArgs e)
        {
            Form fMonitor = Application.OpenForms["frmMonitor"];
            if (fMonitor == null)
            {
                fMonitor = new frmMonitor();
                fMonitor.Show();
            }
            else
            {
                fMonitor.Show();
                fMonitor.BringToFront();
            }
        }

        private static void emails_Click(object sender, EventArgs e)
        {
            Form fEmails = Application.OpenForms["frmEmails"];
            if (fEmails == null)
            {
                fEmails = new frmEmails();
                fEmails.Show();
            }
            else
            {
                fEmails.Show();
                fEmails.BringToFront();
            }
        }

        private static void setupIcon()
        {
            ntIcon.Icon = new System.Drawing.Icon(@"C:\Users\rafas\Desktop\SLN MAILBOT\MailManager\Resources\IconMail.ico");
            ntIcon.Visible = false;
            ntIcon.ContextMenu = new ContextMenu();
            ntIcon.ContextMenu.MenuItems.Add(new MenuItem("Monitor", monitor_Click));
            ntIcon.ContextMenu.MenuItems.Add(new MenuItem("Emails", emails_Click));
            ntIcon.ContextMenu.MenuItems.Add(new MenuItem("Sair", sair_Click));
        }
    }
}
