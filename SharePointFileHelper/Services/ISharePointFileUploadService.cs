using System;
using System.Collections.Generic;
using SharePointFileHelper.Models;

namespace SharePointFileHelper.Services
{
    public interface IOldSharePointFileUploadService
    {
        void UploadFilesToSharePoint(Guid affiliateSiteId, string listName, string targetYear, string formType, string program, List<UploadItem> files);
    }
}
