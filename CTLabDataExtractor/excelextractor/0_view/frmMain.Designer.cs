namespace ExcelExtractor._0_view
{
    partial class frmMain
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmMain));
            this.btnExtract = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.txtSource = new System.Windows.Forms.TextBox();
            this.txtDestination = new System.Windows.Forms.TextBox();
            this.ssSystem = new System.Windows.Forms.StatusStrip();
            this.tssStatus = new System.Windows.Forms.ToolStripStatusLabel();
            this.fdbSourceLocation = new System.Windows.Forms.FolderBrowserDialog();
            this.fdbDestinationLocation = new System.Windows.Forms.FolderBrowserDialog();
            this.tmrProcess = new System.Windows.Forms.Timer(this.components);
            this.nudInterval = new System.Windows.Forms.NumericUpDown();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.chkReprocess = new System.Windows.Forms.CheckBox();
            this.lblLastProc = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.btnSourceBrowse = new System.Windows.Forms.Button();
            this.btnDestBrowse = new System.Windows.Forms.Button();
            this.mainMenu1 = new System.Windows.Forms.MainMenu(this.components);
            this.miSetDefaultLastProc = new System.Windows.Forms.MenuItem();
            this.ssSystem.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.nudInterval)).BeginInit();
            this.SuspendLayout();
            // 
            // btnExtract
            // 
            this.btnExtract.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnExtract.Location = new System.Drawing.Point(454, 122);
            this.btnExtract.Name = "btnExtract";
            this.btnExtract.Size = new System.Drawing.Size(69, 46);
            this.btnExtract.TabIndex = 2;
            this.btnExtract.Text = "START";
            this.btnExtract.UseVisualStyleBackColor = true;
            this.btnExtract.Click += new System.EventHandler(this.btnExtract_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.BackColor = System.Drawing.Color.Transparent;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(12, 5);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(60, 20);
            this.label1.TabIndex = 1;
            this.label1.Text = "Source";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.BackColor = System.Drawing.Color.Transparent;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(12, 56);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(90, 20);
            this.label2.TabIndex = 1;
            this.label2.Text = "Destination";
            // 
            // txtSource
            // 
            this.txtSource.Enabled = false;
            this.txtSource.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtSource.Location = new System.Drawing.Point(16, 29);
            this.txtSource.Name = "txtSource";
            this.txtSource.Size = new System.Drawing.Size(434, 22);
            this.txtSource.TabIndex = 0;
            // 
            // txtDestination
            // 
            this.txtDestination.Enabled = false;
            this.txtDestination.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtDestination.Location = new System.Drawing.Point(16, 79);
            this.txtDestination.Name = "txtDestination";
            this.txtDestination.Size = new System.Drawing.Size(434, 22);
            this.txtDestination.TabIndex = 1;
            // 
            // ssSystem
            // 
            this.ssSystem.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.tssStatus});
            this.ssSystem.Location = new System.Drawing.Point(0, 205);
            this.ssSystem.Name = "ssSystem";
            this.ssSystem.Size = new System.Drawing.Size(537, 22);
            this.ssSystem.TabIndex = 3;
            this.ssSystem.Text = "statusStrip1";
            // 
            // tssStatus
            // 
            this.tssStatus.AutoSize = false;
            this.tssStatus.Name = "tssStatus";
            this.tssStatus.Size = new System.Drawing.Size(320, 17);
            this.tssStatus.Text = "Status:";
            this.tssStatus.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // tmrProcess
            // 
            this.tmrProcess.Interval = 1200000;
            this.tmrProcess.Tick += new System.EventHandler(this.tmrProcess_Tick);
            // 
            // nudInterval
            // 
            this.nudInterval.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.nudInterval.Location = new System.Drawing.Point(161, 135);
            this.nudInterval.Name = "nudInterval";
            this.nudInterval.Size = new System.Drawing.Size(53, 22);
            this.nudInterval.TabIndex = 4;
            this.nudInterval.Value = new decimal(new int[] {
            20,
            0,
            0,
            0});
            this.nudInterval.ValueChanged += new System.EventHandler(this.nudInterval_ValueChanged);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.BackColor = System.Drawing.Color.Transparent;
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.Location = new System.Drawing.Point(214, 139);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(29, 16);
            this.label4.TabIndex = 1;
            this.label4.Text = "Min";
            // 
            // label5
            // 
            this.label5.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.label5.Location = new System.Drawing.Point(0, 111);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(540, 2);
            this.label5.TabIndex = 7;
            // 
            // chkReprocess
            // 
            this.chkReprocess.AutoSize = true;
            this.chkReprocess.BackColor = System.Drawing.Color.Transparent;
            this.chkReprocess.Checked = true;
            this.chkReprocess.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkReprocess.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chkReprocess.Location = new System.Drawing.Point(15, 136);
            this.chkReprocess.Name = "chkReprocess";
            this.chkReprocess.Size = new System.Drawing.Size(140, 20);
            this.chkReprocess.TabIndex = 3;
            this.chkReprocess.Text = "Reprocess Interval";
            this.chkReprocess.UseVisualStyleBackColor = false;
            this.chkReprocess.CheckedChanged += new System.EventHandler(this.chkReprocess_CheckedChanged);
            // 
            // lblLastProc
            // 
            this.lblLastProc.AutoSize = true;
            this.lblLastProc.BackColor = System.Drawing.Color.Transparent;
            this.lblLastProc.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblLastProc.ForeColor = System.Drawing.Color.Green;
            this.lblLastProc.Location = new System.Drawing.Point(1, 182);
            this.lblLastProc.Name = "lblLastProc";
            this.lblLastProc.Size = new System.Drawing.Size(67, 16);
            this.lblLastProc.TabIndex = 9;
            this.lblLastProc.Text = "Last Proc:";
            // 
            // label6
            // 
            this.label6.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.label6.Location = new System.Drawing.Point(0, 177);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(540, 2);
            this.label6.TabIndex = 10;
            // 
            // btnSourceBrowse
            // 
            this.btnSourceBrowse.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnSourceBrowse.Location = new System.Drawing.Point(454, 29);
            this.btnSourceBrowse.Name = "btnSourceBrowse";
            this.btnSourceBrowse.Size = new System.Drawing.Size(69, 23);
            this.btnSourceBrowse.TabIndex = 0;
            this.btnSourceBrowse.Text = "Browse";
            this.btnSourceBrowse.UseVisualStyleBackColor = true;
            this.btnSourceBrowse.Click += new System.EventHandler(this.btnSourceBrowse_Click);
            // 
            // btnDestBrowse
            // 
            this.btnDestBrowse.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnDestBrowse.Location = new System.Drawing.Point(454, 78);
            this.btnDestBrowse.Name = "btnDestBrowse";
            this.btnDestBrowse.Size = new System.Drawing.Size(69, 23);
            this.btnDestBrowse.TabIndex = 1;
            this.btnDestBrowse.Text = "Browse";
            this.btnDestBrowse.UseVisualStyleBackColor = true;
            this.btnDestBrowse.Click += new System.EventHandler(this.btnDestBrowse_Click);
            // 
            // mainMenu1
            // 
            this.mainMenu1.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
            this.miSetDefaultLastProc});
            // 
            // miSetDefaultLastProc
            // 
            this.miSetDefaultLastProc.Index = 0;
            this.miSetDefaultLastProc.Text = "Set Default";
            this.miSetDefaultLastProc.Click += new System.EventHandler(this.miSetDefaultLastProc_Click);
            // 
            // frmMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("$this.BackgroundImage")));
            this.ClientSize = new System.Drawing.Size(537, 227);
            this.Controls.Add(this.btnDestBrowse);
            this.Controls.Add(this.btnSourceBrowse);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.lblLastProc);
            this.Controls.Add(this.chkReprocess);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.nudInterval);
            this.Controls.Add(this.ssSystem);
            this.Controls.Add(this.txtDestination);
            this.Controls.Add(this.txtSource);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnExtract);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Menu = this.mainMenu1;
            this.Name = "frmMain";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "CTLabDataExtractor - V1.1";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.frmMain_FormClosing);
            this.Load += new System.EventHandler(this.frmMain_Load);
            this.ssSystem.ResumeLayout(false);
            this.ssSystem.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.nudInterval)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnExtract;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtSource;
        private System.Windows.Forms.TextBox txtDestination;
        private System.Windows.Forms.StatusStrip ssSystem;
        private System.Windows.Forms.ToolStripStatusLabel tssStatus;
        private System.Windows.Forms.FolderBrowserDialog fdbSourceLocation;
        private System.Windows.Forms.FolderBrowserDialog fdbDestinationLocation;
        private System.Windows.Forms.Timer tmrProcess;
        private System.Windows.Forms.NumericUpDown nudInterval;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.CheckBox chkReprocess;
        private System.Windows.Forms.Label lblLastProc;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Button btnSourceBrowse;
        private System.Windows.Forms.Button btnDestBrowse;
        private System.Windows.Forms.MainMenu mainMenu1;
        private System.Windows.Forms.MenuItem miSetDefaultLastProc;
    }
}

