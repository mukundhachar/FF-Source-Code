using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SlackApiAccess
{
    public class Channels
    {
        [JsonProperty("name")]
        public string name { get; set; }
        [JsonProperty("id")]
        public string id { get; set; }

    }

    public class Messages
    {
        [JsonProperty("type")]
        public string type { get; set; }
        [JsonProperty("user")]
        public string user { get; set; }
        [JsonProperty("text")]
        public string text { get; set; }

        [JsonProperty("Edited")]
        public List<Edited> Edited { get; set; }
        [JsonProperty("ts")]
        public string ts { get; set; }
        [JsonProperty("subtype")]
        public string subtype { get; set; }

        [JsonProperty("SlackFiles")]
        public List<SlackFiles> SlackFiles { get; set; }
        [JsonProperty("attachments")]
        public List<Attachments> attachments { get; set; }

        [JsonProperty("item")]
        public List<PinnedItem> pinned_item { get; set; }
    }

    public class Edited
    {
        [JsonProperty("user")]
        public string user { get; set; }
        [JsonProperty("ts")]
        public string ts { get; set; }

    }
    public class SlackFiles
    {
        [JsonProperty("url_private_download")]
        public string url_private_download { get; set; }


    }
    public class Attachments
    {
        [JsonProperty("text")]
        public string text { get; set; }
        [JsonProperty("thumb_url")]
        public string thumb_url { get; set; }


    }

    public class PinnedItem
    {
        [JsonProperty("user")]
        public string user { get; set; }
        [JsonProperty("comment")]
        public string comment { get; set; }
        [JsonProperty("ts")]
        public string ts { get; set; }


    }
}
