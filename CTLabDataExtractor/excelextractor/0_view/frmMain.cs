using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using System.IO;
using System.Threading;

namespace ExcelExtractor._0_view
{
    public enum FolderLocationType { Source, Destination};

    public partial class frmMain : Form
    {
        ExcelExtractor._1_controller.MainController master = ExcelExtractor._1_controller.MainController.GetInstance();

        public frmMain()
        {
            InitializeComponent();
        }

        #region System Events

            private void frmMain_Load(object sender, EventArgs e)
            {
                tssStatus.Size = new System.Drawing.Size(ssSystem.Size.Width * 9 / 10, tssStatus.Size.Height);
                master.updateStatusDel = new ExcelExtractor._1_controller.UpdateStatusToUI(UpdateStatus);
                master.updateLastProcDel = new ExcelExtractor._1_controller.UpdateLastProc(UpdateProcTimeStamp);
            }

            private void btnExtract_Click(object sender, EventArgs e)
            {
                if (btnExtract.Text.Equals(Constants.BTN_START))
                {
                    if (!IsMandotaryFieldsFilled())
                        MessageBox.Show(Constants.SYS_MANDOTARY_INFO, Constants.SYS_MANDOTARY_FIELD_NOT_FILLED, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    else
                    {
                        btnExtract.Text = Constants.BTN_STOP;
                        tmrProcess.Enabled = chkReprocess.Checked;
                        master.UpdateDirectorySetting(txtSource.Text, txtDestination.Text);
                        EnableUIControls(false);
                        master.Execute(true);
                    }
                }
                else
                {
                    btnExtract.Text = Constants.BTN_START;
                    tmrProcess.Enabled = false;
                    EnableUIControls(true);
                    master.Execute(false);
                    //Thread.Sleep(1000);
                }
            }

            private void btnSourceBrowse_Click(object sender, EventArgs e)
            {
                SelectFolderLocation(FolderLocationType.Source);
            }

            private void btnDestBrowse_Click(object sender, EventArgs e)
            {
                SelectFolderLocation(FolderLocationType.Destination);
            }

            private void chkReprocess_CheckedChanged(object sender, EventArgs e)
            {
                tmrProcess.Enabled = chkReprocess.Checked;
                nudInterval.Enabled = chkReprocess.Checked;
            }

            private void nudInterval_ValueChanged(object sender, EventArgs e)
            {
                tmrProcess.Interval = Convert.ToInt32(nudInterval.Value) * 60 * 1000;
            }

            private void tmrProcess_Tick(object sender, EventArgs e)
            {
                master.Execute(true);
            }

            private void frmMain_FormClosing(object sender, FormClosingEventArgs e)
            {
                master.Execute(false);
                //Thread.Sleep(1000);
            }

            private void miSetDefaultLastProc_Click(object sender, EventArgs e)
            {
                master.SetDefaultLastProcDate();
                MessageBox.Show(Constants.SYS_SET_DEFAULT_LAST_PROC, Constants.SYS_SET_DEFAULT_LAST_PROC_HEADER, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

        #endregion

        #region Private Methods

            private bool IsMandotaryFieldsFilled()
            {
                return !(string.IsNullOrEmpty(txtSource.Text) || string.IsNullOrEmpty(txtDestination.Text));
            }

            private void EnableUIControls(bool flag)
            {
                btnSourceBrowse.Enabled = flag;
                btnDestBrowse.Enabled = flag;
                chkReprocess.Enabled = flag;
                nudInterval.Enabled = flag;
                miSetDefaultLastProc.Enabled = flag;
            }

            private void SelectFolderLocation(FolderLocationType type)
            {
                bool isValidPath = true;

                try
                {
                    string originalPath = Path.GetFullPath(type == FolderLocationType.Source ? txtSource.Text : txtDestination.Text);
                }
                catch
                {
                    isValidPath = false;
                }

                if (isValidPath)
                {
                    if(type == FolderLocationType.Source)
                        fdbSourceLocation.SelectedPath = txtSource.Text;
                    else
                        fdbDestinationLocation.SelectedPath = txtDestination.Text;
                }

                DialogResult result = type == FolderLocationType.Source ? fdbSourceLocation.ShowDialog() : fdbDestinationLocation.ShowDialog();
                if (result == System.Windows.Forms.DialogResult.OK)
                {
                    if(type == FolderLocationType.Source)
                        txtSource.Text = fdbSourceLocation.SelectedPath;
                    else
                        txtDestination.Text = fdbDestinationLocation.SelectedPath;
                }
            }

        #endregion

        #region Delegate Method

            private void UpdateStatus(string msg)
            {
                tssStatus.Text = msg;
            }

            private void UpdateProcTimeStamp()
            {
                lblLastProc.Text =  Constants.SYS_LAST_PROC_PREFIX + DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToShortTimeString();
            }

        #endregion
    }
}
