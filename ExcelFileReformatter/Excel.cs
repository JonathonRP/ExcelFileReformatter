using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Utilities;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace ExcelFileReformatter
{
    public class Excel
    {
        private _Application excel;
        private Workbook workbook;
        private Sheets worksheets;
        private AsyncOperation asyncOperation;

        public ProgressChangedEventHandler FileProgressChanged;
        public RunWorkerCompletedEventHandler FileCompleted;

        private bool CancellationPending { get; set; }

        public Excel()
        {
            asyncOperation = AsyncOperationManager.CreateOperation(null);
        }

        public void Open(string path)
        {
            excel = new Application();
            workbook = excel.Workbooks.Open(path);
            worksheets = workbook.Worksheets;

            excel.Visible = false;
            excel.DisplayAlerts = false;
        }

        private void OnFileProgressChanged(ProgressChangedEventArgs e)
        {
            FileProgressChanged?.Invoke(this, e);
        }

        private void FileProgressReporter(object arg)
        {
            OnFileProgressChanged((ProgressChangedEventArgs)arg);
        }

        private void ReportFileProgress(int percentProgress, object userState = null)
        {
            asyncOperation.Post(FileProgressReporter, new ProgressChangedEventArgs(percentProgress, userState));
        }

        public void Cancel()
        {
            CancellationPending = true;
        }

        protected virtual void OnCompleted(RunWorkerCompletedEventArgs e)
        {
            FileCompleted?.Invoke(this, e);
        }

        private void OperationCompleted(object arg)
        {
            CancellationPending = false;
            OnCompleted((RunWorkerCompletedEventArgs)arg);
        }

        private void RemoveRows(int startIndex, int endIndex, CancelEventArgs e)
        {
            foreach (Worksheet worksheet in worksheets)
            {
                List<Range> removeRows = new List<Range>();
                foreach (Range worksheetRow in worksheet.Rows)
                {
                    if (CancellationPending)
                    {
                        e.Cancel = true;
                        return;
                    }

                    int row = worksheetRow.Row;

                    if (row >= startIndex && row <= endIndex)
                    {
                        removeRows.Add(worksheetRow.EntireRow);

                        int percent = (row / endIndex) * 100;
                        ReportFileProgress(percent, Status.File);
                    }
                    else if (row > endIndex)
                    {
                        Range first = removeRows.First();
                        Range last = removeRows.Last();

                        Range removeRange = worksheet.Range[first, last];
                        removeRange.Delete();

                        return;
                    }
                }
            }
        }

        public void ReformatTo(string path, int startIndex, int endIndex)
        {
            bool cancelled = false;
            object result = (object)null;
            Exception error = (Exception)null;

            try
            {
                DoWorkEventArgs e = new DoWorkEventArgs(null);
                RemoveRows(startIndex, endIndex, e);

                if (e.Cancel)
                    cancelled = true;
                else
                {
                    result = e.Result;
                    SaveAs(path);
                }
            }
            catch (Exception ex)
            {
                error = ex;
            }
            finally
            {
                Close();
            }

            asyncOperation.Post(OperationCompleted, (object)new RunWorkerCompletedEventArgs(result, error, cancelled));
        }

        private void SaveAs(string path)
        {
            workbook.SaveAs(path);
        }

        private void Close()
        {
            workbook.Close(0);
            excel.Quit();

            Marshal.FinalReleaseComObject(worksheets);
            Marshal.FinalReleaseComObject(workbook);
            Marshal.FinalReleaseComObject(excel);
        }
    }
}
