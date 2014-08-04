using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using System.ComponentModel;
using System.IO;
using System.Globalization;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Core;
using System.Runtime.InteropServices;

namespace ExcelExtractor._1_controller
{
    public delegate void UpdateStatusToUI(string msg);
    public delegate void UpdateLastProc();

    class MainController
    {
        private static MainController _instance;
        private ExcelExtractor._2_model.DirectoryModel directory = ExcelExtractor._2_model.DirectoryModel.GetInstance();
        BackgroundWorker worker = new BackgroundWorker();

        public UpdateStatusToUI updateStatusDel;
        public UpdateLastProc updateLastProcDel;

        private MainController()
        {
            InitWorker();
        }

        public static MainController GetInstance()
        {
            if (_instance == null)
                _instance = new MainController();

            return _instance;
        }

        #region Private Methods

            private void InitWorker()
            {
                worker.DoWork += new DoWorkEventHandler(worker_DoWork);
                worker.ProgressChanged += new ProgressChangedEventHandler
                        (worker_ProgressChanged);
                worker.RunWorkerCompleted += new RunWorkerCompletedEventHandler
                        (worker_RunWorkerCompleted);
                worker.WorkerReportsProgress = true;
                worker.WorkerSupportsCancellation = true;
            }

            void worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
            {
                if (e.Cancelled)
                {
                    updateStatusDel(Constants.PROC_CANCEL);
                    EventHandler.Log(Constants.PROC_CANCEL);
                }

                else if (e.Error != null)
                {
                    updateStatusDel(Constants.PROC_ERROR_OCCURRED);
                    EventHandler.Log(Constants.PROC_ERROR_OCCURRED);
                }
                else
                {
                    updateStatusDel(Constants.PROC_COMPLETED);
                    EventHandler.Log(Constants.PROC_COMPLETED);
                }

                updateLastProcDel();
            }

            void worker_ProgressChanged(object sender, ProgressChangedEventArgs e)
            {
                string msg = e.UserState as string;
                if (msg.Contains(";"))
                {
                    List<string> msgList = msg.Split(';').ToList();
                    EventHandler.Log(msg);
                    updateStatusDel(msg);
                }
                else if (msg.Equals(Constants.SYS_CLEAR_CLIPBOARD))
                {
                    System.Windows.Forms.Clipboard.Clear();
                }
                else
                {
                    EventHandler.Log(msg);
                    updateStatusDel(msg);
                }
            }

            void worker_DoWork(object sender, DoWorkEventArgs e)
            {
                BackgroundWorker wk = (BackgroundWorker)sender;
                ExcelExtractor._2_model.DirectoryModel dir = (ExcelExtractor._2_model.DirectoryModel)e.Argument;
                ExtractionController extractor = ExtractionController.GetInstance();
                try
                {
                    extractor.StartExtraction(ref wk, ref e, ref dir);
                }
                catch (Exception ex)
                {
                    wk.ReportProgress(0, ex.Message);
                }
            }

        #endregion

        #region Public Methods

            public void Execute(bool start)
            {
                try
                {
                    if (start && !worker.IsBusy)
                        worker.RunWorkerAsync(directory);
                    else
                    {
                        if (worker.IsBusy)
                            worker.CancelAsync();
                        else
                        {
                            updateStatusDel(Constants.PROC_CANCEL);
                            EventHandler.Log(Constants.PROC_CANCEL);
                        }
                    }
                }
                catch(Exception ex)
                {
                    List<string> exceptionMsg = ex.Message.Split(';').ToList();
                    updateStatusDel(exceptionMsg[0]);
                    EventHandler.Log(exceptionMsg[1] + " " + exceptionMsg[2]);
                }
            }

            public void UpdateDirectorySetting(string original, string dest)
            {
                directory.Source = original;
                directory.Destination = dest;
            }

            public void SetDefaultLastProcDate()
            {
                LastProcController lastProc = LastProcController.GetInstance();
                lastProc.WriteLastProc(Constants.SYS_LAST_PROC_INI_VALUE);
            }

        #endregion
    }
}
