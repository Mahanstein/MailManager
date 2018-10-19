using System;
using System.IO;
using System.Windows.Forms;

namespace MailManager.Logging
{
    public static class Logger
    {
        public static string Path { get; private set;  }
        public static string FileFormat { get; private set; }

        static Logger()
        {
            Path = Application.StartupPath + @"\Log\";
            FileFormat = string.Format("Log_{0:yyyyMMdd}.txt",DateTime.Today);
        }

        public static void Log(string message)
        {
            Directory.CreateDirectory(Path);
            using (StreamWriter writer = File.AppendText(Path + FileFormat))
            {
                string pre = string.Format("[{0:dd/MM/yyyy}][{0:HH:mm:ss}][{1}] - ", DateTime.Now, Environment.UserName);
                writer.WriteLine(pre + message);
            }
        }

        public static void SetPath(string pPath)
        {
            Path = pPath;
        }

        public static void SetFileFormat(string pFormat)
        {
            FileFormat = pFormat;
        }
    }
}
