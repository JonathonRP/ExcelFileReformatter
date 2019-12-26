using System.Windows.Forms;

namespace ExcelFileReformatter
{
    partial class ExcelFileFormatter
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ExcelFileFormatter));
            this.statusStrip1 = new System.Windows.Forms.StatusStrip();
            this.lblStatusFolder = new System.Windows.Forms.ToolStripStatusLabel();
            this.pbFolder = new System.Windows.Forms.ToolStripProgressBar();
            this.lblStatusFile = new System.Windows.Forms.ToolStripStatusLabel();
            this.pbFile = new System.Windows.Forms.ToolStripProgressBar();
            this.lblStatusNotification = new System.Windows.Forms.ToolStripStatusLabel();
            this.tvFoldersFiles = new System.Windows.Forms.TreeView();
            this.imageList1 = new System.Windows.Forms.ImageList(this.components);
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.tsmiFile = new System.Windows.Forms.ToolStripMenuItem();
            this.tssmiAdd = new System.Windows.Forms.ToolStripMenuItem();
            this.tssmiSetRowSkipRange = new System.Windows.Forms.ToolStripMenuItem();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.lblSource = new System.Windows.Forms.Label();
            this.lblDestination = new System.Windows.Forms.Label();
            this.btnSource = new System.Windows.Forms.Button();
            this.btnDestination = new System.Windows.Forms.Button();
            this.pnlSource = new System.Windows.Forms.Panel();
            this.pnlDestination = new System.Windows.Forms.Panel();
            this.panel3 = new System.Windows.Forms.Panel();
            this.btnFormat = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.contextMenuStrip1 = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.ctsmiRemove = new System.Windows.Forms.ToolStripMenuItem();
            this.ctsmiSetRowSkipRange = new System.Windows.Forms.ToolStripMenuItem();
            this.backgroundWorker = new System.ComponentModel.BackgroundWorker();
            this.statusStrip1.SuspendLayout();
            this.menuStrip1.SuspendLayout();
            this.pnlSource.SuspendLayout();
            this.pnlDestination.SuspendLayout();
            this.panel3.SuspendLayout();
            this.contextMenuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // statusStrip1
            // 
            this.statusStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.lblStatusFolder,
            this.pbFolder,
            this.lblStatusFile,
            this.pbFile,
            this.lblStatusNotification});
            this.statusStrip1.Location = new System.Drawing.Point(0, 428);
            this.statusStrip1.Name = "statusStrip1";
            this.statusStrip1.Size = new System.Drawing.Size(637, 22);
            this.statusStrip1.SizingGrip = false;
            this.statusStrip1.TabIndex = 0;
            this.statusStrip1.Text = "statusStrip1";
            // 
            // lblStatusFolder
            // 
            this.lblStatusFolder.Name = "lblStatusFolder";
            this.lblStatusFolder.Size = new System.Drawing.Size(46, 17);
            this.lblStatusFolder.Text = "Folder :";
            this.lblStatusFolder.ToolTipText = "Folder Status";
            // 
            // pbFolder
            // 
            this.pbFolder.Name = "pbFolder";
            this.pbFolder.Size = new System.Drawing.Size(90, 16);
            // 
            // lblStatusFile
            // 
            this.lblStatusFile.Name = "lblStatusFile";
            this.lblStatusFile.Size = new System.Drawing.Size(31, 17);
            this.lblStatusFile.Text = "File :";
            this.lblStatusFile.ToolTipText = "File Status";
            // 
            // pbFile
            // 
            this.pbFile.Name = "pbFile";
            this.pbFile.Size = new System.Drawing.Size(90, 16);
            // 
            // lblStatusNotification
            // 
            this.lblStatusNotification.Font = new System.Drawing.Font("Segoe UI", 8F);
            this.lblStatusNotification.Name = "lblStatusNotification";
            this.lblStatusNotification.Size = new System.Drawing.Size(228, 17);
            this.lblStatusNotification.Text = "New File Created and Location Notification";
            this.lblStatusNotification.ToolTipText = "New File Created and Location Notification";
            // 
            // tvFoldersFiles
            // 
            this.tvFoldersFiles.ImageIndex = 2;
            this.tvFoldersFiles.ImageList = this.imageList1;
            this.tvFoldersFiles.Location = new System.Drawing.Point(12, 97);
            this.tvFoldersFiles.Name = "tvFoldersFiles";
            this.tvFoldersFiles.SelectedImageIndex = 2;
            this.tvFoldersFiles.ShowLines = false;
            this.tvFoldersFiles.ShowPlusMinus = false;
            this.tvFoldersFiles.ShowRootLines = false;
            this.tvFoldersFiles.Size = new System.Drawing.Size(614, 288);
            this.tvFoldersFiles.TabIndex = 3;
            this.tvFoldersFiles.AfterCollapse += new System.Windows.Forms.TreeViewEventHandler(this.TvFoldersFiles_AfterCollapse);
            this.tvFoldersFiles.AfterExpand += new System.Windows.Forms.TreeViewEventHandler(this.TvFoldersFiles_AfterExpand);
            this.tvFoldersFiles.NodeMouseClick += new System.Windows.Forms.TreeNodeMouseClickEventHandler(this.TvFoldersFiles_NodeMouseClick);
            // 
            // imageList1
            // 
            this.imageList1.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList1.ImageStream")));
            this.imageList1.TransparentColor = System.Drawing.Color.Transparent;
            this.imageList1.Images.SetKeyName(0, "icons8-folder-16.png");
            this.imageList1.Images.SetKeyName(1, "icons8-live-folder-16.png");
            this.imageList1.Images.SetKeyName(2, "icons8-xls-16.png");
            // 
            // menuStrip1
            // 
            this.menuStrip1.BackColor = System.Drawing.SystemColors.Control;
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.tsmiFile});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(637, 24);
            this.menuStrip1.TabIndex = 5;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // tsmiFile
            // 
            this.tsmiFile.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.tssmiAdd,
            this.tssmiSetRowSkipRange});
            this.tsmiFile.Name = "tsmiFile";
            this.tsmiFile.Size = new System.Drawing.Size(37, 20);
            this.tsmiFile.Text = "File";
            // 
            // tssmiAdd
            // 
            this.tssmiAdd.Name = "tssmiAdd";
            this.tssmiAdd.Size = new System.Drawing.Size(160, 22);
            this.tssmiAdd.Text = "Add";
            this.tssmiAdd.ToolTipText = "Add a File to be formatted";
            this.tssmiAdd.Click += new System.EventHandler(this.TssmiAdd_Click);
            // 
            // tssmiSetRowSkipRange
            // 
            this.tssmiSetRowSkipRange.Name = "tssmiSetRowSkipRange";
            this.tssmiSetRowSkipRange.Size = new System.Drawing.Size(160, 22);
            this.tssmiSetRowSkipRange.Text = "Set Rows to Skip";
            this.tssmiSetRowSkipRange.Click += new System.EventHandler(this.tssmiSetRowSkipRange_Click);
            // 
            // lblSource
            // 
            this.lblSource.AutoEllipsis = true;
            this.lblSource.AutoSize = true;
            this.lblSource.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.lblSource.Location = new System.Drawing.Point(3, 6);
            this.lblSource.Margin = new System.Windows.Forms.Padding(3);
            this.lblSource.Name = "lblSource";
            this.lblSource.Size = new System.Drawing.Size(73, 13);
            this.lblSource.TabIndex = 6;
            this.lblSource.Text = "Source Folder";
            this.lblSource.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // lblDestination
            // 
            this.lblDestination.AutoEllipsis = true;
            this.lblDestination.AutoSize = true;
            this.lblDestination.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.lblDestination.Location = new System.Drawing.Point(3, 6);
            this.lblDestination.Name = "lblDestination";
            this.lblDestination.Size = new System.Drawing.Size(92, 13);
            this.lblDestination.TabIndex = 7;
            this.lblDestination.Text = "Destination Folder";
            this.lblDestination.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // btnSource
            // 
            this.btnSource.FlatAppearance.BorderColor = System.Drawing.Color.Black;
            this.btnSource.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnSource.Location = new System.Drawing.Point(82, 1);
            this.btnSource.Name = "btnSource";
            this.btnSource.Size = new System.Drawing.Size(75, 23);
            this.btnSource.TabIndex = 8;
            this.btnSource.Text = "Browse";
            this.btnSource.UseVisualStyleBackColor = true;
            this.btnSource.Click += new System.EventHandler(this.BtnSource_Click);
            // 
            // btnDestination
            // 
            this.btnDestination.FlatAppearance.BorderColor = System.Drawing.Color.Black;
            this.btnDestination.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnDestination.Location = new System.Drawing.Point(101, 1);
            this.btnDestination.Name = "btnDestination";
            this.btnDestination.Size = new System.Drawing.Size(75, 23);
            this.btnDestination.TabIndex = 9;
            this.btnDestination.Text = "Browse";
            this.btnDestination.UseVisualStyleBackColor = true;
            this.btnDestination.Click += new System.EventHandler(this.BtnDestination_Click);
            // 
            // pnlSource
            // 
            this.pnlSource.AllowDrop = true;
            this.pnlSource.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.pnlSource.Controls.Add(this.btnSource);
            this.pnlSource.Controls.Add(this.lblSource);
            this.pnlSource.Location = new System.Drawing.Point(0, 7);
            this.pnlSource.Name = "pnlSource";
            this.pnlSource.Size = new System.Drawing.Size(614, 25);
            this.pnlSource.TabIndex = 10;
            this.pnlSource.DragDrop += new System.Windows.Forms.DragEventHandler(this.PnlSource_DragDrop);
            this.pnlSource.DragEnter += new System.Windows.Forms.DragEventHandler(this.pnlSource_DragEnter);
            this.pnlSource.DragOver += new System.Windows.Forms.DragEventHandler(this.pnlSource_DragOver);
            this.pnlSource.DragLeave += new System.EventHandler(this.pnlSource_DragLeave);
            // 
            // pnlDestination
            // 
            this.pnlDestination.AllowDrop = true;
            this.pnlDestination.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.pnlDestination.Controls.Add(this.btnDestination);
            this.pnlDestination.Controls.Add(this.lblDestination);
            this.pnlDestination.Location = new System.Drawing.Point(0, 33);
            this.pnlDestination.Name = "pnlDestination";
            this.pnlDestination.Size = new System.Drawing.Size(614, 25);
            this.pnlDestination.TabIndex = 11;
            this.pnlDestination.DragDrop += new System.Windows.Forms.DragEventHandler(this.PnlDestination_DragDrop);
            this.pnlDestination.DragEnter += new System.Windows.Forms.DragEventHandler(this.pnlDestination_DragEnter);
            this.pnlDestination.DragOver += new System.Windows.Forms.DragEventHandler(this.pnlDestination_DragOver);
            this.pnlDestination.DragLeave += new System.EventHandler(this.pnlDestination_DragLeave);
            // 
            // panel3
            // 
            this.panel3.Controls.Add(this.pnlDestination);
            this.panel3.Controls.Add(this.pnlSource);
            this.panel3.Location = new System.Drawing.Point(12, 27);
            this.panel3.Name = "panel3";
            this.panel3.Size = new System.Drawing.Size(614, 64);
            this.panel3.TabIndex = 12;
            // 
            // btnFormat
            // 
            this.btnFormat.Location = new System.Drawing.Point(551, 395);
            this.btnFormat.Name = "btnFormat";
            this.btnFormat.Size = new System.Drawing.Size(75, 25);
            this.btnFormat.TabIndex = 13;
            this.btnFormat.Text = "Format";
            this.btnFormat.UseVisualStyleBackColor = true;
            this.btnFormat.Click += new System.EventHandler(this.btnFormat_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(470, 395);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 25);
            this.btnCancel.TabIndex = 14;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            this.openFileDialog1.Multiselect = true;
            // 
            // contextMenuStrip1
            // 
            this.contextMenuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.ctsmiRemove,
            this.ctsmiSetRowSkipRange});
            this.contextMenuStrip1.Name = "contextMenuStrip1";
            this.contextMenuStrip1.ShowImageMargin = false;
            this.contextMenuStrip1.Size = new System.Drawing.Size(136, 48);
            // 
            // ctsmiRemove
            // 
            this.ctsmiRemove.Name = "ctsmiRemove";
            this.ctsmiRemove.Size = new System.Drawing.Size(135, 22);
            this.ctsmiRemove.Text = "Remove";
            // 
            // ctsmiSetRowSkipRange
            // 
            this.ctsmiSetRowSkipRange.Name = "ctsmiSetRowSkipRange";
            this.ctsmiSetRowSkipRange.Size = new System.Drawing.Size(135, 22);
            this.ctsmiSetRowSkipRange.Text = "Set Rows to Skip";
            // 
            // backgroundWorker
            // 
            this.backgroundWorker.WorkerReportsProgress = true;
            this.backgroundWorker.WorkerSupportsCancellation = true;
            this.backgroundWorker.DoWork += new System.ComponentModel.DoWorkEventHandler(this.backgroundWorker_DoWork);
            this.backgroundWorker.ProgressChanged += new System.ComponentModel.ProgressChangedEventHandler(this.backgroundWorker_ProgressChanged);
            this.backgroundWorker.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.backgroundWorker_RunWorkerCompleted);
            // 
            // ExcelFileFormatter
            // 
            this.AllowDrop = true;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(637, 450);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnFormat);
            this.Controls.Add(this.panel3);
            this.Controls.Add(this.tvFoldersFiles);
            this.Controls.Add(this.statusStrip1);
            this.Controls.Add(this.menuStrip1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MainMenuStrip = this.menuStrip1;
            this.MaximizeBox = false;
            this.Name = "ExcelFileFormatter";
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide;
            this.Text = "ExcelFileFormatter";
            this.DragDrop += new System.Windows.Forms.DragEventHandler(this.ExcelFileFormatter_DragDrop);
            this.DragEnter += new System.Windows.Forms.DragEventHandler(this.ExcelFileFormatter_DragEnter);
            this.statusStrip1.ResumeLayout(false);
            this.statusStrip1.PerformLayout();
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.pnlSource.ResumeLayout(false);
            this.pnlSource.PerformLayout();
            this.pnlDestination.ResumeLayout(false);
            this.pnlDestination.PerformLayout();
            this.panel3.ResumeLayout(false);
            this.contextMenuStrip1.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.StatusStrip statusStrip1;
        private System.Windows.Forms.ToolStripStatusLabel lblStatusFolder;
        private System.Windows.Forms.ToolStripProgressBar pbFolder;
        private System.Windows.Forms.ToolStripStatusLabel lblStatusFile;
        private System.Windows.Forms.ToolStripProgressBar pbFile;
        private System.Windows.Forms.TreeView tvFoldersFiles;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem tsmiFile;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
        private System.Windows.Forms.Label lblSource;
        private System.Windows.Forms.Label lblDestination;
        private System.Windows.Forms.Button btnSource;
        private System.Windows.Forms.Button btnDestination;
        private System.Windows.Forms.Panel pnlSource;
        private System.Windows.Forms.Panel pnlDestination;
        private System.Windows.Forms.Panel panel3;
        private System.Windows.Forms.ToolStripStatusLabel lblStatusNotification;
        private System.Windows.Forms.Button btnFormat;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.ImageList imageList1;
        private ContextMenuStrip contextMenuStrip1;
        private ToolStripMenuItem ctsmiRemove;
        private ToolStripMenuItem ctsmiSetRowSkipRange;
        private ToolStripMenuItem tssmiAdd;
        private ToolStripMenuItem tssmiSetRowSkipRange;
        private System.ComponentModel.BackgroundWorker backgroundWorker;
    }
}

