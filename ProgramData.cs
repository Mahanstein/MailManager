using MailManager.MailService;
using System.Collections.Generic;

namespace MailManager
{
    static class ProgramData
    {
        public static string LogPath;
        public static string EmailBackupPath;

        public static List<MonitoredMailbox> MailboxList = new List<MonitoredMailbox>();
    }
}
