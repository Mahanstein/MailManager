using MailManager.MailService;
using System;
using System.ComponentModel;
using System.Windows.Forms;

namespace MailManager
{
    public partial class frmMonitor : Form
    {
        private BackgroundWorker _worker;

        public frmMonitor()
        {
            InitializeComponent();
            init();
        }

        private void init()
        {
            Logging.Logger.Log("Abrindo form Monitor.");
            startWorker();
        }

        private void startWorker()
        {
            setupWorker();
            _worker.RunWorkerAsync();
        }

        private void setupWorker()
        {
            if (_worker == null)
            {
                _worker = new BackgroundWorker();
                _worker.WorkerReportsProgress = true;
                _worker.ProgressChanged += Worker_ProgressChanged;
                _worker.RunWorkerCompleted += _worker_RunWorkerCompleted;
                _worker.DoWork += _worker_DoWork;
            }
        }

        private void _worker_DoWork(object sender, DoWorkEventArgs e)
        {
            startProcess();
        }

        private void _worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            display("Fim do trabalho assíncrono.");
        }

        private void Worker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            display(e.UserState.ToString());
        }

        private void startProcess()
        {
            try
            {
                Logging.Logger.Log("Iniciando processo.");
                _worker.ReportProgress(0, "Iniciando processo.");

                MailService.Service oMailService = new Service();
                oMailService.MonitoredMailboxes.Add(new MonitoredMailbox("rafas.rafaelsilva@hotmail.com", "TesteManagerIn", "TesteManagerProcessed"));

                _worker.ReportProgress(0, "Conectando...");
                oMailService.Connect("rafas.rafaelsilva@hotmail.com", "tom4ess4");

                Logging.Logger.Log("Copiando emails.");
                foreach (MonitoredMailbox item in oMailService.MonitoredMailboxes)
                {
                    try
                    {
                        _worker.ReportProgress(0, item.OrginalFolderName + " -> " + item.DestinationFolderName);
                        _worker.ReportProgress(0, oMailService.CountEmails(item.OrginalFolderName).ToString() + " emails.");

                        foreach (string message in oMailService.CopyEmails(item.OriginalFolder, item.DestinationFolder))
                        {
                            _worker.ReportProgress(0, "Copiando email: " + message);
                        }
                    }
                    catch (Exception ex)
                    {
                        _worker.ReportProgress(0, "Erro geral na cópia: " + ex.Message);
                        Logging.Logger.Log("Erro geral na cópia: " + ex.Message);
                    }
                }

                Logging.Logger.Log("Importando emails.");
                foreach (var item in oMailService.MonitoredMailboxes)
                {
                    try
                    {
                        foreach (string message in oMailService.DownloadEmails(item.OriginalFolder))
                        {
                            _worker.ReportProgress(0, "Baixando email: " + message);
                        }
                    }
                    catch (Exception ex)
                    {
                        _worker.ReportProgress(0, "Erro geral na cópia local: " + ex.Message);
                        Logging.Logger.Log("Erro geral na cópia local: " + ex.Message);
                    }
                }
            }
            catch (Exception ex)
            {
                _worker.ReportProgress(0, "Erro geral no processo: " + ex.Message);
                Logging.Logger.Log("Erro geral na cópia: " + ex.Message);
            }
        }

        private void display(string pMessage)
        {
            lvMonitor.Items.Add(pMessage);
            lvMonitor.Columns[0].Width = -1;
        }

        private void frmMonitor_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing)
            {
                Hide();
                Logging.Logger.Log("Fechando monitor.");
                e.Cancel = true;
            }
        }
    }
}
