using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using System.Collections.Specialized;
using System.Net;
using Newtonsoft.Json.Linq;
using System.Threading;
using System.Configuration;
using SlackApiAccess;
public class SlackClient
{
	private readonly Uri uri;
	private readonly Encoding encoding = new UTF8Encoding();
    private string StrLegacyToken;
    private string SlackAPIChannelLstUrl;
    private string SlackAPIUserLstUrl;
    private string SlackAPIChannelHistoryMsgUrl;

    public SlackClient(string LegacyToken)
	{
		//uri = new Uri(urlWithAccessToken);
        StrLegacyToken = LegacyToken;
        SlackAPIChannelLstUrl = Utility.GetConfigValueByKey("SlackAPIChannelLstUrl");
        SlackAPIUserLstUrl = Utility.GetConfigValueByKey("SlackAPIUserLstUrl");
        SlackAPIChannelHistoryMsgUrl = Utility.GetConfigValueByKey("SlackAPIChannelHistoryMsgUrl");
	}
	
	//Post a message using simple strings
	public void PostMessage(string text, string username = null, string channel = null)
	{
        try
        {     
		    Payload payload = new Payload()
		    {
			    Channel = channel,
			    Username = username,
			    Text = text
		    };
		
		    PostMessage(payload);
        }
        catch (System.Exception ex)
        {

            throw ;
        }
	}
	
	
    /// <summary>
    /// Post a message using a Payload object to slack
    /// </summary>
    /// <param name="payload"></param>
	public void PostMessage(Payload payload)
	{
		string payloadJson = JsonConvert.SerializeObject(payload);		
		using (WebClient client = new WebClient())
		{
			NameValueCollection data = new NameValueCollection();
			data["payload"] = payloadJson;	
			var response = client.UploadValues(uri, "POST", data);        
			//The response text is usually "ok"
			string responseText = encoding.GetString(response);
		}
	}



    /// <summary>
    /// Fetching channel text message from slack
    /// </summary>
    /// <param name="text"></param>
    /// <param name="username"></param>
    /// <param name="channel"></param>
    public void GetMessageFormChannel(string text, string username = null, string channel = null)
    {
        Payload payload = new Payload()
        {
            Channel = channel,
            Username = username,
            Text = text
        };

        GetMessageFormChannel(payload);
    }


  
    /// <summary>
    /// Gettting channel list 
    /// </summary>
    /// <param name="payload"></param>
    public void GetMessageFormChannel(Payload payload)
    {

        string payloadJson = JsonConvert.SerializeObject(payload);
        using (WebClient client = new WebClient())
        {
            Thread.Sleep(TimeSpan.FromSeconds(5));
            var response = client.UploadValues(SlackAPIChannelLstUrl, "POST", new NameValueCollection() {
        {"token",""+StrLegacyToken+""},
        {"channel","#general"},
         {"Retry-After","6"}
            });
                   string responseText = encoding.GetString(response);
             
                    }
    }

  /// <summary>
  /// 
  /// Getting all the channal information form slack
  /// </summary>
  /// <returns></returns>
    public string GetChannelInfoFromSlack()
    {
        using (WebClient client = new WebClient())
        {
            Thread.Sleep(TimeSpan.FromSeconds(5));
            var response = client.UploadValues(SlackAPIChannelLstUrl, "POST", new NameValueCollection() {           
            {"token",""+StrLegacyToken+""},
            {"pretty","1"},
             {"Retry-After","6"}
            });
            return encoding.GetString(response);
        }
    }

    /// <summary>
    /// Gettting Channesl message for the given channel ID
    /// </summary>
    /// <param name="chanellId"></param>
    /// <returns></returns>
    public string GetChannelMessagesInfoFromSlack(string chanellId,double oldest)
    {
        using (WebClient client = new WebClient())
        {
            var response = client.UploadValues(SlackAPIChannelHistoryMsgUrl, "POST", new NameValueCollection() {           
            {"token",""+StrLegacyToken+""},
            {"channel",""+chanellId+""},
            //{"oldest",""+oldest==null?"0":oldest+""},            
             {"inclusive","1"},
             {"pretty","1"},
             {"Retry-After","6"}

                });



            return encoding.GetString(response);
        }
    }



    /// <summary>
    /// Get the User Dispaly name for the given salck user Id
    /// </summary>
    /// <param name="SlackUserId"></param>
    /// <returns></returns>
    public string GetUserDisplayNameFromSlack(String SlackUserId)
    {
        string SlackUserDisplayName = "";
        using (WebClient client = new WebClient())
        {
            var userList = client.UploadValues(SlackAPIUserLstUrl, "POST", new NameValueCollection() {
                {"token",""+StrLegacyToken+""},
                {"Retry-After","6"}});
                    string strUserList = encoding.GetString(userList);
                    JObject o = JObject.Parse(strUserList);
                    JArray arr = (JArray)o.SelectToken("members");
                    var name = arr.FirstOrDefault(x => x.Value<string>("id") == ""+SlackUserId+"").Value<string>("real_name");
                    SlackUserDisplayName = name;
                     }

        return SlackUserDisplayName;

    }

     /// <summary>
    /// Get all slack user list
    /// </summary>
    /// <param name="SlackUserId"></param>
    /// <returns></returns>
    public string GetSlackUserList()
    {
        
        using (WebClient client = new WebClient())
        {
            var userList = client.UploadValues(SlackAPIUserLstUrl, "POST", new NameValueCollection() {
                {"token",""+StrLegacyToken+""},
                {"Retry-After","6"}});
            return encoding.GetString(userList);
        }

    }






    }




//This class serializes into the Json payload required by Slack Incoming WebHooks
public class Payload
{
	[JsonProperty("channel")]
	public string Channel { get; set; }
	
	[JsonProperty("username")]
	public string Username { get; set; }
	
	[JsonProperty("text")]
	public string Text { get; set; }
}


