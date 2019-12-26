using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;

namespace ExcelFileReformatter
{
    public class File : ISetRowSkipRange
    {
        private readonly System.IO.FileInfo _FileInfo;
        public FileInfo FileInfo { get; }

        public Tuple<int, int> RowSkipRange { get; private set; }

        public File(string fileName)
        {
            FileInfo = new FileInfo(fileName);
            _FileInfo = FileInfo;
        }

        public void SetRowSkipRange(int start, int end)
        {
            RowSkipRange = new Tuple<int, int>(start, end);
        }

        public static FileAttributes GetAttributes(string path)
        {
            return System.IO.File.GetAttributes(path);
        }

        public FileStream Open(FileMode fileMode)
        {
            return _FileInfo.Open(fileMode);
        }
        public FileStream Open(FileMode fileMode, FileAccess fileAccess)
        {
            return _FileInfo.Open(fileMode, fileAccess);
        }
        public FileStream Open(FileMode fileMode, FileAccess fileAccess, FileShare fileShare)
        {
            return _FileInfo.Open(fileMode, fileAccess, fileShare);
        }
        public FileStream Create()
        {
            return _FileInfo.Create();
        }
        public FileStream OpenRead()
        {
            return _FileInfo.OpenRead();
        }
        public FileStream OpenWrite()
        {
            return _FileInfo.OpenWrite();
        }
        public StreamReader OpenText()
        {
            return _FileInfo.OpenText();
        }
    }
}
