using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelFileReformatter
{
    public class Folder : ISetRowSkipRange
    {
        public FolderInfo FolderInfo { get; }
        public List<File> Files { get; }

        public Tuple<int, int> RowSkipRange { get; private set; }

        public Folder()
        {

        }
        public Folder(string path, List<File> files): this()
        {
            FolderInfo = new FolderInfo(path);
            Files = files;
        }

        public void SetRowSkipRange(int start, int end)
        {
            Files.ForEach(f => f.SetRowSkipRange(start, end));
            RowSkipRange = new Tuple<int, int>(start, end);
        }

        public Folder Remove(List<string> fileNames)
        {
            fileNames.ForEach(n => Files.Remove(Files.Find(file => file.FileInfo.Name == n)));
            return this;
        }
    }
}
