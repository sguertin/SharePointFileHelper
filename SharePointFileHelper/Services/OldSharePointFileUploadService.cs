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
    public class OldSharePointFileUploadService : IOldSharePointFileUploadService
    {
        private readonly string _clientId;
        private readonly string _clientSecret;
        private readonly string _siteUrl;

        private readonly ILogger<OldSharePointFileUploadService> _logger;
        
        public OldSharePointFileUploadService(IConfiguration configuration, ILogger<OldSharePointFileUploadService> logger)
        {
            _clientId = configuration["SharePointClientId"];
            _clientSecret = configuration["SharePointClientSecret"];
            _siteUrl = configuration["SharePointSiteUrl"];
            _logger = logger;
        }
        public void UploadFilesToSharePoint(Guid affiliateSiteId, string listName, string targetYear, string formType, string program, List<UploadItem> files)
        {
            var authManager = new AuthenticationManager();

            var targetUrl = $"{_siteUrl}/{affiliateSiteId}";

            using var ctx = authManager.GetACSAppOnlyContext(targetUrl, _clientId, _clientSecret);

            var web = ctx.Web;
            ctx.Load(web);
            ctx.Load(web.Lists);
            
            ctx.ExecuteQueryRetry();
            
            var list = web.Lists.GetByTitle(listName);
            ctx.Load(list);
            ctx.Load(list.ContentTypes);
            ctx.ExecuteQueryRetry();


            var insuranceDocSetTypeId = list.ContentTypes
                .Where(c => c.Name == "Insurance Doc Set")
                .Select(c => c.Id)
                .FirstOrDefault();
            var insuranceDocumentTypeId = list.ContentTypes
                .Where(c => c.Name == "Insurance")
                .Select(c => c.Id)
                .FirstOrDefault();

            var folder = list.RootFolder.EnsureFolder(targetYear);

            ctx.Load(folder);
            ctx.Load(folder.ListItemAllFields);
            ctx.Load(folder.ListItemAllFields.ContentType);
            ctx.ExecuteQueryRetry();
            
            var folderToUpload = web.GetFolderByServerRelativeUrl(folder.ServerRelativeUrl);
            if (folderToUpload == null)
            {
                throw new DirectoryNotFoundException($"Failed to create directory {targetYear} in {targetUrl}");
            }
            
            ctx.Load(folderToUpload);
            ctx.Load(folderToUpload.ListItemAllFields);
            ctx.ExecuteQueryRetry();

            if (insuranceDocSetTypeId != null &&
                folderToUpload.ListItemAllFields["ContentTypeId"] != insuranceDocSetTypeId)
            {
                folderToUpload.ListItemAllFields["ContentTypeId"] = insuranceDocSetTypeId;
                folderToUpload.ListItemAllFields.Update();
            }
            folderToUpload.Update();
            ctx.Load(folderToUpload);
            ctx.ExecuteQuery();

            foreach (var file in files)
            {
                folderToUpload.UploadFile(file.FileName, new MemoryStream(file.FileContent), true);
                folderToUpload.Update();
                ctx.ExecuteQueryRetry();

                folderToUpload.EnsureProperty(f => f.ServerRelativeUrl);
                
                var serverRelativeUrl = folderToUpload.ServerRelativeUrl.TrimEnd('/') + '/' + file.FileName;
                var uploadedFile = web.GetFileByServerRelativeUrl(serverRelativeUrl);
                
                ctx.Load(uploadedFile);
                ctx.Load(uploadedFile.ListItemAllFields);
                ctx.Load(uploadedFile.ListItemAllFields.ContentType);
                ctx.ExecuteQueryRetry();
                if (insuranceDocumentTypeId != null)
                {
                    uploadedFile.ListItemAllFields["ContentTypeId"] = insuranceDocumentTypeId.ToString();
                    uploadedFile.ListItemAllFields["Insurance_x0020_Doc_x0020_Type"] = formType;
                    uploadedFile.ListItemAllFields["Program"] = program;
                    uploadedFile.ListItemAllFields.Update();
                }
                uploadedFile.Update();
                ctx.ExecuteQueryRetry();

                if (!uploadedFile.Exists)
                {
                    _logger.LogWarning($"Could not find newly uploaded file: {file.FileName} at {serverRelativeUrl}");
                }
            }
        }

    }
}
