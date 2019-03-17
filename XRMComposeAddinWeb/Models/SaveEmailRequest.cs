using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using Newtonsoft.Json;

namespace XRMComposeAddinWeb.Models
{
    public class SaveEmailRequest
    {
        [JsonProperty(PropertyName ="Title")]
        public string Subject { get; set; }
        public string Message { get; set; }
        public string From { get; set; }
        public string To { get; set; }
        public string InOut { get; set; }
        [JsonProperty(PropertyName = "CategoryLookupId")]
        public string Category { get; set; }
        public string RelatedItemListId { get; set; }
        public DateTime Received { get; set; }
        public string ConversationId { get; set; }
        public string ConversationTopic { get; set; }
        public string RelatedItemId { get; set; }
    }
}