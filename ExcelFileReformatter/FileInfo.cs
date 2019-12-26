using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelFileReformatter
{
    public class FileInfo
    {
        private System.IO.FileInfo _FileInfo { get; }

        public string Name => _FileInfo.Name;
        public DirectoryInfo Directory => _FileInfo.Directory;
        public string DirectoryName => _FileInfo.DirectoryName;
        public string Extension => _FileInfo.Extension;
        public string FullName => _FileInfo.FullName;
        public bool Exists => _FileInfo.Exists;
        public bool IsReadOnly => _FileInfo.IsReadOnly;
        public long Length => _FileInfo.Length;
        public FileAttributes Attributes => _FileInfo.Attributes;
        public DateTime CreationTime => _FileInfo.CreationTime;
        public DateTime CreationTimeUtc => _FileInfo.CreationTimeUtc;
        public DateTime LastAccessTime => _FileInfo.LastAccessTime;
        public DateTime LastAccessTimeUtc => _FileInfo.LastAccessTimeUtc;
        public DateTime LastWriteTime => _FileInfo.LastWriteTime;
        public DateTime LastWriteTimeUtc => _FileInfo.LastWriteTimeUtc;

        public FileInfo(string fileName)
        {
            _FileInfo = new System.IO.FileInfo(fileName);
        }

        public static implicit operator System.IO.FileInfo(FileInfo fileInfo)
        {
             return new System.IO.FileInfo(fileInfo.FullName);
        }
    }
}
