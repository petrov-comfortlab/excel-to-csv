using System;
using System.IO;

namespace ExcelToCsv
{
    public class File
    {
        public File(string fullPath) => FullPath = fullPath;
        public string FullPath { get; private set; }

        public string FileName
        {
            get => Path.GetFileName(FullPath);
            set => FullPath = Path.Combine(Path.GetDirectoryName(FullPath) ?? throw new InvalidOperationException(), value);
        }
    }
}