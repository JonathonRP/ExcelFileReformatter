using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Utilities;

namespace ExcelFileReformatter
{
    public partial class ExcelFileFormatter : Form
    {
        private Folder Folder { get; set; }
        private List<File> Files { get; }
        private List<File> AddedFiles { get; set; }
        private List<Folder> AddedFolders { get; set; }

        private Excel Excel { get; set; }

        private Action<TreeNodeMouseClickEventArgs> ShowContextMenu;
        private Action<TreeNode> Remove;
        private EventHandler Removal;
        private Action<TreeNode> SetItemRowSkipRange;
        private EventHandler ShowSetItemRowSkipRange;
        private Action ShowSetRowSkipRange;

        private Action<DragEventArgs> FolderDragEnter;
        private Action<DragEventArgs> ItemDragEnter;

        private SetRowSkipRange SetRowSkipRangeForm { get; set; }

        public ExcelFileFormatter()
        {
            Folder = new Folder();
            AddedFiles = new List<File>();
            AddedFolders = new List<Folder>();
            Files = new List<File>();

            ShowContextMenu = OnShowContextMenu;
            Remove = OnRemove;
            SetItemRowSkipRange = OnSetItemRowSkipRange;
            ShowSetRowSkipRange = OnShowSetRowSkipRange;

            FolderDragEnter = OnFolderDragEnter;
            ItemDragEnter = OnItemDragEnter;

            Icon = Properties.Resources.ExcelFileFormatter;
            InitializeComponent();
        }

        private bool OpenFilesDialog(out List<string> files)
        {
            openFileDialog1.FileName = "";
            openFileDialog1.Filter = "Excel File|*.xls;*.xlsx;*.xls*";

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                files = openFileDialog1.FileNames.ToList();
                return true;
            }

            files = null;
            return false;
        }
        private DialogResult OpenFolderDialog(out string folder)
        {
            DialogResult result = folderBrowserDialog1.ShowDialog();
            if (result == DialogResult.OK)
            {
                folder = folderBrowserDialog1.SelectedPath;
                return result;
            }

            folder = null;
            return result;
        }

        private List<File> ExtractFiles(string folder)
        {
            return Directory.GetFiles(folder, "*.xls*").Select(ProcessFile).ToList();
        }

        private TreeNode ProcessFolder(Folder folder)
        {
            return new TreeNode(folder.FolderInfo.Name, folder.Files.Select(ProcessFile).ToArray()) { ImageIndex = 0, SelectedImageIndex = 0 };
        }
        private Folder ProcessFolder(string folder)
        {
            return new Folder(Path.GetFileName(folder), ExtractFiles(folder));
        }

        private List<Folder> ProcessFolders(List<string> folders)
        {
            return folders.Select(ProcessFolder).ToList();
        }

        private TreeNode ProcessFile(File file)
        {
            return new TreeNode(file.FileInfo.Name);
        }
        private File ProcessFile(string fileName)
        {
            return new File(fileName); ;
        }

        private List<File> ProcessFiles(List<string> fileNames)
        {
            return fileNames.Select(ProcessFile).ToList();
        }

        private void TssmiAdd_Click(object sender, EventArgs e)
        {
            if (!OpenFilesDialog(out List<string> fileNames)) return;
            
            AddedFiles = ProcessFiles(fileNames);
            Files.AddRange(AddedFiles);

            tvFoldersFiles.Nodes.AddRange(AddedFiles.Select(ProcessFile).ToArray());
        }

        private void tssmiSetRowSkipRange_Click(object sender, EventArgs e)
        {
            if (Folder.Files == null)
            {
                MessageBox.Show("Please select a source folder with Excel files first", "Setting Row Skip Range Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            SetRowSkipRangeForm = new SetRowSkipRange(Folder);
            ShowSetRowSkipRange();
        }

        private void BtnSource_Click(object sender, EventArgs e)
        {
            if (OpenFolderDialog(out string folder) == DialogResult.Cancel) return;
            ProcessSourceFolder(folder);
        }

        private void ProcessSourceFolder(string folderName)
        {
            var folder = ProcessFolder(folderName);

            if (folder.Files.Count < 1)
            {
                MessageBox.Show("Folder doesn't contain Excel files", "Opening Folder Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (Folder.Files != null)
            {
                foreach (File file in Folder.Files)
                {
                    Files.Remove(file);
                }

                tvFoldersFiles.Nodes.Remove(tvFoldersFiles.Nodes.Cast<TreeNode>().First(n => n.Text == Folder.FolderInfo.Name));
            }

            lblSource.Text = $"Source Folder: {folderName}";
            btnSource.Left = lblSource.Width + 10;

            Folder = folder;
            Files.AddRange(folder.Files);
            tvFoldersFiles.Nodes.Add(ProcessFolder(Folder));
        }

        private void BtnDestination_Click(object sender, EventArgs e)
        {
            if (OpenFolderDialog(out string folder) == DialogResult.Cancel) return;
            ProcessDestinationFolder(folder);
        }

        private void ProcessDestinationFolder(string folder)
        {
            lblDestination.Text = $"Destination Folder: {folder}";
            btnDestination.Left = lblDestination.Width + 10;
        }

        private void pnlSource_DragEnter(object sender, DragEventArgs e)
        {
            Highlight(sender,true);
            FolderDragEnter(e);
        }

        private void pnlSource_DragOver(object sender, DragEventArgs e)
        {
            Highlight(sender, true);
        }

        private void pnlSource_DragLeave(object sender, EventArgs e)
        {
            Highlight(sender, false);
        }

        private void PnlSource_DragDrop(object sender, DragEventArgs e)
        {
            string[] folders = (string[])e.Data.GetData(DataFormats.FileDrop);

            if (folders.Length == 1)
            {
                ProcessSourceFolder(folders[0]);
            }
            else
            {
                MessageBox.Show("Please select only 1 folder as source.", "Too Many Folders Selected", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

            Highlight(sender, false);
        }

        private void pnlDestination_DragEnter(object sender, DragEventArgs e)
        {
            Highlight(sender, true);
            FolderDragEnter(e);
        }

        private void pnlDestination_DragOver(object sender, DragEventArgs e)
        {
            Highlight(sender, true);
        }

        private void PnlDestination_DragDrop(object sender, DragEventArgs e)
        {
            string[] folders = (string[])e.Data.GetData(DataFormats.FileDrop);

            if (folders.Length == 1)
            {
                ProcessDestinationFolder(folders[0]);
            }
            else
            {
                MessageBox.Show("Please select only 1 folder for destination.", "Too Many Folders Selected", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

            Highlight(sender, false);
        }

        private void pnlDestination_DragLeave(object sender, EventArgs e)
        {
            Highlight(sender, false);
        }

        private Label GetLabel(Panel panel)
        {
            return panel.Controls.Cast<Control>().ToList().Find(c => c.GetType() == typeof(Label)) as Label;
        }

        private void Highlight(object sender, bool highlight)
        {
            if (sender.GetType() == typeof(Panel))
            {
                GetLabel(sender as Panel).Font = highlight ? new Font(DefaultFont, FontStyle.Underline) : DefaultFont;
            }
            else
            {
                (sender as Control).BackColor = highlight ? ControlPaint.Dark(DefaultBackColor) : DefaultBackColor;
            }
        }

        private void ExcelFileFormatter_DragDrop(object sender, DragEventArgs e)
        {
            List<string> items = ((string[])e.Data.GetData(DataFormats.FileDrop)).ToList();
            List<string> folderList = new List<string>();
            List<string> fileList = new List<string>();

            items.ForEach(item =>
            {
                if (File.GetAttributes(item).HasFlag(FileAttributes.Directory))
                {
                    folderList.Add(item);
                }

                if (Path.GetExtension(item).Contains(".xls"))
                {
                    fileList.Add(item);
                }
            });

            AddedFolders = ProcessFolders(folderList);
            Files.AddRange(AddedFolders.SelectMany(folder => folder.Files));
            tvFoldersFiles.Nodes.AddRange(AddedFolders.Select(ProcessFolder).ToArray());

            AddedFiles = ProcessFiles(fileList);
            Files.AddRange(AddedFiles);
            tvFoldersFiles.Nodes.AddRange(AddedFiles.Select(ProcessFile).ToArray());
        }

        private void ExcelFileFormatter_DragEnter(object sender, DragEventArgs e)
        {
            ItemDragEnter(e);
        }

        private void TvFoldersFiles_NodeMouseClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            switch (e.Button)
            {
                case MouseButtons.Left when e.Node.Parent != null:
                    return;
                case MouseButtons.Left when e.Node.IsExpanded:
                    e.Node.Collapse();
                    break;
                case MouseButtons.Left:
                    e.Node.Expand();
                    break;
                case MouseButtons.Right:
                    ShowContextMenu(e);
                    break;
            }
        }

        private void TvFoldersFiles_AfterCollapse(object sender, TreeViewEventArgs e)
        {
            e.Node.SelectedImageIndex = 0;
        }

        private void TvFoldersFiles_AfterExpand(object sender, TreeViewEventArgs e)
        {
            e.Node.SelectedImageIndex = 1;
        }

        private void OnShowContextMenu(TreeNodeMouseClickEventArgs e)
        {
            e.Node.ContextMenuStrip = contextMenuStrip1;

            ctsmiRemove.Click += Removal = (sender, args) => 
            { 
                Remove(e.Node); 
                ctsmiRemove.Click -= Removal; 
                ctsmiSetRowSkipRange.Click -= ShowSetItemRowSkipRange;
            };

            ctsmiSetRowSkipRange.Click += ShowSetItemRowSkipRange = (sender, args) =>
            {
                SetItemRowSkipRange(e.Node);
                ctsmiRemove.Click -= Removal;
                ctsmiSetRowSkipRange.Click -= ShowSetItemRowSkipRange;
            };
        }
        private void OnRemove(TreeNode e)
        {
            if (e.ImageIndex == -1)
            {
                Files.Remove(Files.Find(file => file.FileInfo.Name == e.Text));

                if (e.Parent != null)
                {
                    List<File> removedFiles = Folder.Files.Where(file => !Files.Contains(file)).ToList();
                    Folder.Remove(removedFiles.Select(f => f.FileInfo.Name).ToList());
                }
                else
                {
                    AddedFiles.Remove(new File(e.Text));
                }
            }
            else
            {
                if (e.Text == Folder.FolderInfo?.Name)
                {
                    List<File> removedFiles = Folder.Files.Select(file =>
                    {
                        if (!Files.Contains(file)) return null;
                        Files.Remove(file);
                        return file;
                    }).ToList();

                    Folder.Remove(removedFiles.Select(f => f.FileInfo.Name).ToList());
                    Folder = new Folder();
                    lblSource.Text = "Source Folder";
                    btnSource.Left = lblSource.Width + 10;
                }
                else
                {
                    List<File> removedFiles = AddedFolders.SelectMany(folder => folder.Files.Select(file =>
                    {
                        if (!Files.Contains(file)) return null;
                        Files.Remove(file);
                        return file;
                    })).ToList();

                    AddedFolders.Remove(AddedFolders.Find(folder => folder.FolderInfo.Name == e.Text).Remove(removedFiles.Select(f => f.FileInfo.Name).ToList()));
                }
            }

            tvFoldersFiles.Nodes.Remove(e);
        }
        private void OnSetItemRowSkipRange(TreeNode e)
        {
            SetRowSkipRangeForm = e.ImageIndex == -1 ? new SetRowSkipRange(Files.Find(file => file.FileInfo.Name == e.Text)) : 
                e.Text == Folder.FolderInfo?.Name ? new SetRowSkipRange(Folder) : 
                new SetRowSkipRange(AddedFolders.Find(folder => folder.FolderInfo.Name == e.Text));

            ShowSetRowSkipRange();
        }

        private void OnShowSetRowSkipRange()
        {
            SetRowSkipRangeForm.StartPosition = FormStartPosition.CenterParent;
            SetRowSkipRangeForm.ShowDialog(this);
        }

        private void OnFolderDragEnter(DragEventArgs e)
        {
            var paths = (string[])e.Data.GetData(DataFormats.FileDrop);
            e.Effect = paths.Any(path => !File.GetAttributes(path).HasFlag(FileAttributes.Directory)) ? DragDropEffects.None : DragDropEffects.Copy;
        }

        private void OnItemDragEnter(DragEventArgs e)
        {
            var paths = (string[])e.Data.GetData(DataFormats.FileDrop);
            e.Effect = paths.Any(path => (File.GetAttributes(path).HasFlag(FileAttributes.Directory) && !(ExtractFiles(path).Count > 0)) || (File.GetAttributes(path).HasFlag(FileAttributes.Normal) && !(Path.GetExtension(path).Contains(".xls"))))
                ? DragDropEffects.None : DragDropEffects.Copy;
        }

        private void btnFormat_Click(object sender, EventArgs e)
        {
            if (Files == null || Files.Count <= 0 || Files.TrueForAll(f => f.RowSkipRange == null))
            {
                MessageBox.Show("Please set Row Skip Range", "Row Skip Range Not Set", MessageBoxButtons.OK, MessageBoxIcon.Error);
                tssmiSetRowSkipRange.PerformClick();
                return;
            }

            if (lblDestination.Text == "Destination Folder")
            {
                MessageBox.Show("Please select destination folder", "Formatting Files Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                btnDestination.PerformClick();
                return;
            }

            if (!backgroundWorker.IsBusy)
            {
                Excel = new Excel()
                {
                    FileProgressChanged = backgroundWorker_ProgressChanged
                };

                backgroundWorker.RunWorkerAsync();
                backgroundWorker.ReportProgress(0, Status.Reset);
            }
            else
            {
                MessageBox.Show("Sorry currently busy working in the background. Press cancel, if you wish to cancel current operation.", "Formatting Files Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            Excel.Cancel();
            backgroundWorker.CancelAsync();
        }

        private void backgroundWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            Files.ForEach(f =>
            {
                void Cancel()
                {
                    e.Cancel = true;
                }

                if (backgroundWorker.CancellationPending)
                {
                    Cancel();
                    return;
                }

                backgroundWorker.ReportProgress(0, Status.File);

                File newFile = new File($"{lblDestination.Text.Replace("Destination Folder: ", "")}\\{f.FileInfo.Name}");

                Excel.Open(f.FileInfo.FullName);
                Excel.ReformatTo(newFile.FileInfo.FullName, f.RowSkipRange.Item1, f.RowSkipRange.Item2);

                Excel.FileCompleted = OnCompleted;

                void OnCompleted (object controller, RunWorkerCompletedEventArgs eventArgs)
                {
                    if (eventArgs.Cancelled)
                    {
                        Cancel();
                    }

                    if (eventArgs.Error != null)
                    {
                        throw eventArgs.Error;
                    }

                    if (!eventArgs.Cancelled && eventArgs.Error == null)
                    {
                        decimal index = Files.FindIndex(file => file.FileInfo.Name == newFile.FileInfo.Name) + 1;
                        decimal percentDecimal = index / Folder.Files.Count;
                        decimal percent = percentDecimal * 100;

                        backgroundWorker.ReportProgress((int)percent, newFile);
                    }
                }
            });
        }

        private void backgroundWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            switch (e.UserState)
            {
                case Status state:
                    switch (state)
                    {
                        case Status.Folder:
                            pbFolder.Value = e.ProgressPercentage;
                            break;
                        case Status.File:
                            pbFile.Value = e.ProgressPercentage;
                            break;
                        case Status.Reset:
                            pbFolder.Value = e.ProgressPercentage;
                            pbFile.Value = e.ProgressPercentage;
                            break;
                        default:
                            throw new ArgumentOutOfRangeException();
                    }
                    break;
                case File newFile:
                    pbFolder.Value = e.ProgressPercentage;
                    lblStatusNotification.ForeColor = Color.Black;
                    lblStatusNotification.Text = $"Created : {newFile.FileInfo.Directory.Name}\\{newFile.FileInfo.Name}";
                    break;
            }
        }

        private void backgroundWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Cancelled)
            {
                pbFolder.Value = 0;
                pbFile.Value = 0;
                lblStatusNotification.ForeColor = Color.Red;
                lblStatusNotification.Text = "Cancelled...";
            }
            if(e.Error != null)
            {
                MessageBox.Show($"{e.Error.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            if (!e.Cancelled && e.Error == null)
            {
                MessageBox.Show("Completed formatting", "Completed", MessageBoxButtons.OK, MessageBoxIcon.None);
            }
        }
    }
}
