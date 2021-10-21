using System;
using System.Collections.Generic;
using SharePointFileHelper.Models;

namespace SharePointFileHelper.Services
{
    public interface ISharePointLoaderService
    {
        /// <summary>
        /// Upload a file to a List on a SharePoint site
        /// </summary>
        /// <param name="siteName">The SharePoint Site with the List</param>
        /// <param name="listName">The List to upload the file to</param>
        /// <param name="file">The information about the file</param>
        /// <param name="destinationFolderPath">The folder path to upload to (Defaults to RootFolder)</param>
        /// <returns>The url to the newly uploaded file</returns>
        void UploadFileToSharePoint(string siteName, string listName, UploadFile file, string destinationFolderPath = "/");
    }
}
