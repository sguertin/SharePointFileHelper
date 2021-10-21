using System.Collections.Generic;
namespace SharePointFileHelper.Models
{
    public abstract class SharePointUploadItem
    {
        public string Name { get; set; }
        public string ContentType { get; set; }
        public IDictionary<string, object> ItemFieldData { get; set; } = new Dictionary<string,object>();
    }
}