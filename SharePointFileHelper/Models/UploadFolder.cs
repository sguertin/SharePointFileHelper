using System;
using System.Collections.Generic;

namespace SharePointFileHelper.Models
{
    public class UploadFolder : SharePointUploadItem
    { 
        public string FolderPath { get; set; } = "/";
        public List<UploadFile> Files { get; set; }        
    }
}