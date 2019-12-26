using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelFileReformatter
{
    public class FolderInfo
    {
        private DirectoryInfo _DirectoryInfo { get; }

        public string Name => _DirectoryInfo.Name;
        public DirectoryInfo Parent => _DirectoryInfo.Parent;
        public DirectoryInfo Root => _DirectoryInfo.Root;
        public string Extension => _DirectoryInfo.Extension;
        public string FullName => _DirectoryInfo.FullName;
        public bool Exists => _DirectoryInfo.Exists;
        public FileAttributes Attributes => _DirectoryInfo.Attributes;
        public DateTime CreationTime => _DirectoryInfo.CreationTime;
        public DateTime CreationTimeUtc => _DirectoryInfo.CreationTimeUtc;
        public DateTime LastAccessTime => _DirectoryInfo.LastAccessTime;
        public DateTime LastAccessTimeUtc => _DirectoryInfo.LastAccessTimeUtc;
        public DateTime LastWriteTime => _DirectoryInfo.LastWriteTime;
        public DateTime LastWriteTimeUtc => _DirectoryInfo.LastWriteTimeUtc;

        public FolderInfo(string path)
        {
            _DirectoryInfo = new DirectoryInfo(path);
        }

        public static implicit operator DirectoryInfo(FolderInfo folderInfo)
        {
            return new DirectoryInfo(folderInfo.FullName);
        }
    }
}
