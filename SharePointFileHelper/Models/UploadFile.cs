using System;
using System.IO;
using System.Collections.Generic;

namespace SharePointFileHelper.Models
{
    public class UploadFile : SharePointUploadItem
    { 
        public string FileSourcePath { get; set; }
        public byte[] FileContent { get; set; } = null;
        public byte[] GetFileContents()
        {
            if (FileContent == null)
            { 
                FileContent = File.ReadAllBytes(FileSourcePath);
            }
            return FileContent;
        }
        public string FileName => Path.GetFileName(FileSourcePath);
    }
}