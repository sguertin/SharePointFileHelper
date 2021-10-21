using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.SharePoint.Client;
using PnP.Framework;
using SharePointFileHelper.Models;

namespace SharePointFileHelper.Services
{
    public class SharePointLoaderService : ISharePointLoaderService
    {

        private readonly string _clientId;
        private readonly string _clientSecret;
        private readonly string _siteUrl;
        private readonly ILogger<SharePointLoaderService> _logger;
        private readonly AuthenticationManager _authManager;

        public SharePointLoaderService(IConfiguration configuration, ILogger<SharePointLoaderService> logger)
        {
            _clientId = configuration["SharePointClientId"];
            _clientSecret = configuration["SharePointClientSecret"];
            _siteUrl = configuration["SharePointSiteUrl"];
            _logger = logger;
            _authManager = new AuthenticationManager();
        }

        public void UploadFileToSharePoint(string siteName, string listName, UploadFile file, string destinationFolderPath = "/")
        {        
            using var ctx = _authManager.GetACSAppOnlyContext(GetTargetUrl(siteName), _clientId, _clientSecret);
            var web = ctx.Web;
            ctx.Load(web);
            ctx.Load(web.Lists);
            
            ctx.ExecuteQueryRetry();
            
            var list = web.Lists.GetByTitle(listName);
            ctx.Load(list);

            var folder = list.RootFolder;
            ctx.Load(folder);
            ctx.ExecuteQueryRetry();

            if (destinationFolderPath != "/")
            {
                var foldersInPath = destinationFolderPath.Split('/');
                foreach (var folderName in foldersInPath)
                {
                    if (string.IsNullOrEmpty(folderName))
                    {
                        continue;
                    }
                    folder = folder.EnsureFolder(folderName);
                    ctx.Load(folder);
                    ctx.ExecuteQueryRetry();
                }
            }
            var folderToUpload = web.GetFolderByServerRelativeUrl(folder.ServerRelativeUrl);
            if (folderToUpload == null)
            {
                throw new DirectoryNotFoundException($"Failed to create directory '{folderToUpload.Name}'");
            }            
            ctx.Load(folderToUpload);
            ctx.Load(folderToUpload.ListItemAllFields);
            ctx.ExecuteQueryRetry();

            folderToUpload.UploadFile(file.Name, new MemoryStream(file.FileContent), true);
            folderToUpload.Update();
            ctx.ExecuteQueryRetry();
            
            folderToUpload.EnsureProperty(f => f.ServerRelativeUrl);

            var serverRelativeUrl = folderToUpload.ServerRelativeUrl.TrimEnd('/') + '/' + file.Name;
            var uploadedFile = web.GetFileByServerRelativeUrl(serverRelativeUrl);
            if (!uploadedFile.Exists)
            {
                throw new Exception($"Could not find file {file.Name} after uploading!");
            }
            if (file.ItemFieldData.Keys.Count == 0 && string.IsNullOrEmpty(file.ContentType))
            {
                return;
            }
            ctx.Load(uploadedFile);
            ctx.Load(uploadedFile.ListItemAllFields);
            ctx.Load(uploadedFile.ListItemAllFields.ContentType);
            ctx.ExecuteQueryRetry();
            if (!string.IsNullOrEmpty(file.ContentType))
            {
                var contentTypeId = GetContentId(siteName, listName, file.ContentType);
                uploadedFile.ListItemAllFields["ContentTypeId"] = contentTypeId.ToString();
            }
            foreach (var key in file.ItemFieldData.Keys)
            {
                var fieldName = key.Replace(" ","_x0020_");
                uploadedFile.ListItemAllFields[fieldName] = file.ItemFieldData[key];
            }
            uploadedFile.ListItemAllFields.Update();
            uploadedFile.Update();
            ctx.ExecuteQueryRetry();
        }
        // TODO: Update with correct return type
        private object GetContentId(string siteName, string listName, string contentTypeName)
        {
            using var ctx = _authManager.GetACSAppOnlyContext(GetTargetUrl(siteName), _clientId, _clientSecret);
            var web = ctx.Web;
            ctx.Load(web);
            ctx.Load(web.Lists);
            
            ctx.ExecuteQueryRetry();
            
            var list = web.Lists.GetByTitle(listName);
            ctx.Load(list);
            ctx.Load(list.ContentTypes);
            ctx.ExecuteQueryRetry();

            return list.ContentTypes
                .Where(c => c.Name == contentTypeName)
                .Select(c => c.Id)
                .FirstOrDefault();
        }

        private string GetTargetUrl(string siteName) => $"{_siteUrl}/{siteName}";
        
        
        
        // public void UploadFilesToSharePoint(Guid affiliateSiteId, string listName, string targetYear, string formType, string program, List<UploadItem> files)
        // {
        //     var authManager = new AuthenticationManager();

        //     var targetUrl = $"{_siteUrl}/{affiliateSiteId}";

        //     using var ctx = authManager.GetACSAppOnlyContext(targetUrl, _clientId, _clientSecret);

        //     var web = ctx.Web;
        //     ctx.Load(web);
        //     ctx.Load(web.Lists);
            
        //     ctx.ExecuteQueryRetry();
            
        //     var list = web.Lists.GetByTitle(listName);
        //     ctx.Load(list);
        //     ctx.Load(list.ContentTypes);
        //     ctx.ExecuteQueryRetry();


        //     var insuranceDocSetTypeId = list.ContentTypes
        //         .Where(c => c.Name == "Insurance Doc Set")
        //         .Select(c => c.Id)
        //         .FirstOrDefault();
        //     var insuranceDocumentTypeId = list.ContentTypes
        //         .Where(c => c.Name == "Insurance")
        //         .Select(c => c.Id)
        //         .FirstOrDefault();

        //     var folder = list.RootFolder.EnsureFolder(targetYear);

        //     ctx.Load(folder);
        //     ctx.Load(folder.ListItemAllFields);
        //     ctx.Load(folder.ListItemAllFields.ContentType);
        //     ctx.ExecuteQueryRetry();
            
        //     var folderToUpload = web.GetFolderByServerRelativeUrl(folder.ServerRelativeUrl);
        //     if (folderToUpload == null)
        //     {
        //         throw new DirectoryNotFoundException($"Failed to create directory {targetYear} in {targetUrl}");
        //     }
            
        //     ctx.Load(folderToUpload);
        //     ctx.Load(folderToUpload.ListItemAllFields);
        //     ctx.ExecuteQueryRetry();

        //     if (insuranceDocSetTypeId != null &&
        //         folderToUpload.ListItemAllFields["ContentTypeId"] != insuranceDocSetTypeId)
        //     {
        //         folderToUpload.ListItemAllFields["ContentTypeId"] = insuranceDocSetTypeId;
        //         folderToUpload.ListItemAllFields.Update();
        //     }
        //     folderToUpload.Update();
        //     ctx.Load(folderToUpload);
        //     ctx.ExecuteQuery();

        //     foreach (var file in files)
        //     {
        //         folderToUpload.UploadFile(file.FileName, new MemoryStream(file.FileContent), true);
        //         folderToUpload.Update();
        //         ctx.ExecuteQueryRetry();

        //         folderToUpload.EnsureProperty(f => f.ServerRelativeUrl);
                
        //         var serverRelativeUrl = folderToUpload.ServerRelativeUrl.TrimEnd('/') + '/' + file.FileName;
        //         var uploadedFile = web.GetFileByServerRelativeUrl(serverRelativeUrl);
                
        //         ctx.Load(uploadedFile);
        //         ctx.Load(uploadedFile.ListItemAllFields);
        //         ctx.Load(uploadedFile.ListItemAllFields.ContentType);
        //         ctx.ExecuteQueryRetry();
        //         if (insuranceDocumentTypeId != null)
        //         {
        //             uploadedFile.ListItemAllFields["ContentTypeId"] = insuranceDocumentTypeId.ToString();
        //             uploadedFile.ListItemAllFields["Insurance_x0020_Doc_x0020_Type"] = formType;
        //             uploadedFile.ListItemAllFields["Program"] = program;
        //             uploadedFile.ListItemAllFields.Update();
        //         }
        //         uploadedFile.Update();
        //         ctx.ExecuteQueryRetry();

        //         if (!uploadedFile.Exists)
        //         {
        //             _logger.LogWarning($"Could not find newly uploaded file: {file.FileName} at {serverRelativeUrl}");
        //         }
        //     }
    }

}
