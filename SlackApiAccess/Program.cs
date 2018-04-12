using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using System.Net;
using System.Runtime.Serialization;
using Newtonsoft.Json.Linq;
using System.Dynamic;
using System.Threading;
using System.IO;
using iTextSharp;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.Text.RegularExpressions;
using Newtonsoft.Json;
using System.Xml.Linq;
using iTextSharp.text.pdf.parser;
namespace SlackApiAccess
{
    class Program
    {
        static void Main(string[] args)
        {
            MergeSlackMessages();
        }        
        /// <summary>
        /// Pulling slack messages  
        /// </summary>
        static void MergeSlackMessages()
        {
            DateTime? Oldest = null;
            StringBuilder strBuilderMain = new StringBuilder();
            double lngOldest = 0;
            strBuilderMain.AppendLine("----------------------------------------------------------");
            strBuilderMain.AppendLine("Process start date:-" + DateTime.Now);
            Logger.LogInformation(strBuilderMain.ToString());
            strBuilderMain.Clear();
            string outputFile = string.Empty;
            string LegacyToken = Utility.GetConfigValueByKey("LegacyToken");
            string FilterEmoji = Utility.GetConfigValueByKey("FltrEmoji");
            string FilterChannels = Utility.GetConfigValueByKey("FltrChannesl");
            string SlackArchivePath = Utility.GetConfigValueByKey("SlackArchivePath");
            string LastRuntime = Utility.GetConfigValueByKey("LastRuntime");
            string FilterSubType = Utility.GetConfigValueByKey("FilterSubType");
            string[] arrSubType = FilterSubType == string.Empty ? null : FilterSubType.Split(new string[] { "," }, StringSplitOptions.None);
            string JsonChannelsStr = string.Empty;
            string[] arrChannels;
            string JsonUserList = string.Empty; 
            string Channel_name= string.Empty;
            string channelId = string.Empty;
            Paragraph title;
            string strCrntDt = "";
            string strCreatedTime = string.Empty;
            string strCreatedDate = string.Empty;
            string[] arrEmojis;
            string UserCreatedDtTime = string.Empty;
            string[] headers;
            string MsgCreatedUserName = string.Empty;
            string UserMessage = string.Empty;
            string designationFile= string.Empty;
            string souceFile= string.Empty;
            string souceFile2 = string.Empty;

            if (!File.Exists(LastRuntime))
            {
                File.Create(LastRuntime);
            }           
            if (Utility.GetConfigValueByKey("msgFromDate") == string.Empty)
            {
                DateTime dt = File.GetLastWriteTime(LastRuntime);
                Oldest = dt;
            }
            else
            {
                Oldest = Convert.ToDateTime(Utility.GetConfigValueByKey("Olddest"));         
            }            
            if (Oldest == null)
            {
                lngOldest = 0;
            }
            else
            {
                lngOldest = Datetimeconversion(Convert.ToString(Oldest));
            }
            try
            {
                SlackClient client = new SlackClient(LegacyToken);
                JsonChannelsStr = client.GetChannelInfoFromSlack();
                System.Xml.XmlDocument ChannelsxmlDocument = JsonConvert.DeserializeXmlNode(JsonChannelsStr, "root");
                XElement ChannelElement = XElement.Parse(ChannelsxmlDocument.InnerXml);
                var ChannelsList = (from xmlElement in ChannelElement.Elements("channels")
                                    select new Channels
                                    {
                                        name = xmlElement.Element("name").Value,
                                        id = xmlElement.Element("id").Value
                                    }).ToList();

                var listJChannelobject = new List<Channels>();
                 arrChannels = FilterChannels == string.Empty ? null : FilterChannels.Split(new string[] { "," }, StringSplitOptions.None);
                if (arrChannels != null)
                {
                    foreach (var s in arrChannels)
                    {
                        var temp = (from chnl in ChannelsList
                            .Where(m => m.name ==s                                 
                            )
                            select chnl).ToList();
                        temp.ForEach(a => listJChannelobject.Add(a));                      
                    }
                }
                else
                {
                    var temp = (from chnl in ChannelsList
                                select chnl).ToList();
                    temp.ForEach(a => listJChannelobject.Add(a));
                }
                Thread.Sleep(TimeSpan.FromSeconds(5));
                JsonUserList = client.GetSlackUserList();
                JObject results = JObject.Parse(JsonChannelsStr);
                foreach (Channels result in listJChannelobject)
                {                   
                     Channel_name = (string)result.name;
                     channelId = (string)result.id;                    
                    Console.WriteLine("Channel name:" + Channel_name);                   
                    Font titleFont = new Font(FontFactory.GetFont("Microsoft Sans Serif", 10, Font.BOLD));
                    titleFont.SetColor(44, 45, 48);
                    Font regularFont = FontFactory.GetFont("Microsoft Sans Serif", 8);                   
                    Font userNameFont = new Font(FontFactory.GetFont("Microsoft Sans Serif", 8, Font.BOLD));
                    userNameFont.SetColor(44, 45, 48);                    
                    title = new Paragraph("#" + Channel_name, titleFont);
                    title.Alignment = Element.ALIGN_LEFT;
                    Thread.Sleep(TimeSpan.FromSeconds(5));                   
                    string channelMessages = client.GetChannelMessagesInfoFromSlack(channelId, lngOldest);                    
                    System.Xml.XmlDocument xmlDocument = JsonConvert.DeserializeXmlNode(channelMessages, "root");
                    XElement element = XElement.Parse(xmlDocument.InnerXml);
                    var messagesList = (from xmlElement in element.Elements("messages")
                                  select new Messages
                                  {
                                      text = xmlElement.Element("text")!= null ?  xmlElement.Element("text").Value :string.Empty,
                                      type = xmlElement.Element("type") != null ? xmlElement.Element("type").Value : string.Empty,
                                      user = xmlElement.Element("user") != null ? xmlElement.Element("user").Value : string.Empty,
                                      ts =    xmlElement.Element("ts")!= null ?  xmlElement.Element("ts").Value : string.Empty,
                                      subtype = xmlElement.Element("subtype") != null ? xmlElement.Element("subtype").Value : string.Empty,
                                      SlackFiles =
                                      xmlElement.Elements("file") != null ?
                                      (from child in xmlElement.Elements("file")
                                       select new SlackFiles
                                   {
                                       url_private_download = child.Element("url_private_download")!=null?child.Element("url_private_download").Value:string.Empty ,

                                   }).ToList<SlackFiles>()
                                                : null,

                                      Edited = xmlElement.Elements("edited") != null ?
                                      (from child in xmlElement.Elements("edited")
                                       select new Edited
                                       {
                                           ts = child.Element("ts") != null ? child.Element("ts").Value : string.Empty,
                                           user = child.Element("user") != null ? child.Element("user").Value : string.Empty,
                                       }).ToList<Edited>()
                                                : null,

                                          attachments = xmlElement.Elements("attachments") != null ?
                                       (from child in xmlElement.Elements("attachments")

                                        select new Attachments
                                        {
                                            text = child.Element("text") !=null? child.Element("text").Value:null,
                                            thumb_url = child.Element("thumb_url") !=null? child.Element("thumb_url").Value:null,
                                        }
                                        ).ToList<Attachments>()
                                                 : null,
                                      pinned_item = xmlElement.Elements("item") != null ?
                                             (from child in xmlElement.Elements("item")

                                              select new PinnedItem
                                              {
                                                  user = child.Element("user") != null ? child.Element("user").Value : null,
                                                  comment = child.Element("comment") != null ? child.Element("comment").Value : null,
                                                  ts = xmlElement.Element("ts") != null ? xmlElement.Element("ts").Value : null,
                                              }
                                              ).ToList<PinnedItem>()
                                                       : null,


                                  }).ToList();
                   
                    JObject objMsgResults = JObject.Parse(channelMessages);
                    var listJobject = new List<Messages>();
                    arrEmojis = FilterEmoji == string.Empty ? null : FilterEmoji.Split(new string[] { "," }, StringSplitOptions.None);
                    if (arrEmojis != null)
                    {
                        foreach (var s in arrEmojis)
                        {                          
                            var temp = (from msg in messagesList
                             .Where(m =>
                                  (Convert.ToDateTime(UnixTimeStampToDateTime(Convert.ToDouble(m.ts))) > Oldest && m.text.Contains(s)) || (m.Edited.Any(x => Convert.ToDateTime(UnixTimeStampToDateTime(Convert.ToDouble(x.ts))) > Oldest)) || (m.pinned_item.Any(y => Convert.ToDateTime(UnixTimeStampToDateTime(Convert.ToDouble(y.ts))) > Oldest))
                             )
                               select msg).ToList();
                            temp.ForEach(a => listJobject.Add(a));
                        }
                    }
                    else
                    {
                        var temp = (from msg in messagesList

                                       .Where(m =>
                                  (Convert.ToDateTime(UnixTimeStampToDateTime(Convert.ToDouble(m.ts))) > Oldest) || (m.Edited.Any(x => Convert.ToDateTime(UnixTimeStampToDateTime(Convert.ToDouble(x.ts))) > Oldest)) || (m.pinned_item.Any(y => Convert.ToDateTime(UnixTimeStampToDateTime(Convert.ToDouble(y.ts))) > Oldest))
                             )

                                     select msg).ToList();
                        temp.ForEach(a => listJobject.Add(a));
                    }
                    if (listJobject.Count > 0)
                    {
                        outputFile = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.System), SlackArchivePath + "TempChannel" + ".pdf");
                        using (FileStream fs = new FileStream(outputFile, File.Exists(outputFile) ? FileMode.Append : FileMode.OpenOrCreate, FileAccess.Write, FileShare.None))
                        {
                            Document doc = new Document(PageSize.LETTER);
                            using (PdfWriter writer = PdfWriter.GetInstance(doc, fs))
                            {
                                doc.Open();
                                doc.Add(title);
                                
                                foreach (Messages msgResult in listJobject)
                                {
                                    try
                                    {
                                        string MsgCreatedByUserName = string.Empty;
                                        if (msgResult.user != null)
                                        {
                                             UserCreatedDtTime = string.Empty;                                          
                                             strCreatedTime = string.Empty;
                                             strCreatedDate = string.Empty;
                                            if (msgResult.Edited.Count > 0)
                                            {
                                                if (msgResult.ts != null)
                                                {
                                                    UserCreatedDtTime = msgResult.Edited.FirstOrDefault().ts.ToString();
                                                }
                                                 strCreatedTime = Convert.ToDateTime(UnixTimeStampToDateTime(Convert.ToDouble(UserCreatedDtTime))).ToString("h:mm tt");
                                                 strCreatedDate = Convert.ToDateTime(UnixTimeStampToDateTime(Convert.ToDouble(UserCreatedDtTime))).ToString("dddd MMMM dd");

                                            }
                                            else
                                            {
                                                if (msgResult.ts != null)
                                                {
                                                    UserCreatedDtTime = msgResult.ts.ToString();
                                                }
                                                 strCreatedTime = Convert.ToDateTime(UnixTimeStampToDateTime(Convert.ToDouble(UserCreatedDtTime))).ToString("h:mm tt");
                                                 strCreatedDate = Convert.ToDateTime(UnixTimeStampToDateTime(Convert.ToDouble(UserCreatedDtTime))).ToString("dddd MMMM dd");
                                            }
                                           
                                            Font georgiaTimeFont = FontFactory.GetFont("Arial", 7f);
                                            georgiaTimeFont.Color = BaseColor.GRAY;
                                            if (strCrntDt != strCreatedDate)
                                            {
                                                headers = new string[] { " ", " ", " " };
                                                //Font font = new Font("Arial", 8, Font.BOLD);
                                                Font font = new Font(FontFactory.GetFont("Microsoft Sans Serif", 8, Font.BOLD));
                                                font.SetColor(44, 45, 48);
                                                var table1 = new PdfPTable(headers.Length) { WidthPercentage = 100 };
                                                table1.SetWidths(GetHeaderWidths(font, headers));
                                                Paragraph cell3Text1;
                                                for (int i = 0; i < headers.Length; ++i)
                                                {
                                                    if (i == 1)
                                                    {
                                                        PdfPCell cell = new PdfPCell(new Phrase(strCreatedDate, font));
                                                        cell.Border = 0;
                                                        cell.PaddingBottom = 10f;
                                                        cell.PaddingRight = 0;
                                                        cell.PaddingLeft = 0;
                                                        cell.PaddingTop = 4f;
                                                        cell.HorizontalAlignment = (Element.ALIGN_CENTER);
                                                        table1.AddCell(cell);
                                                    }
                                                    else
                                                    {
                                                        cell3Text1 = new Paragraph(new Chunk(new iTextSharp.text.pdf.draw.LineSeparator(0.0F, 100.0F, BaseColor.LIGHT_GRAY, Element.ALIGN_LEFT, 1)));
                                                        PdfPCell CellFirstAndLast = new PdfPCell(cell3Text1);
                                                        CellFirstAndLast.Border = 0;
                                                        table1.AddCell(CellFirstAndLast);
                                                    }
                                                }
                                                int cw1, cw2, cw3 = 0;
                                                cw2 = (int)(strCreatedDate.Length * .9);
                                                cw1 = ((int)((100 - cw2) / 2)) + 1;
                                                cw3 = cw1;
                                                int[] firstTablecellwidth = { cw1, cw2, cw3 };
                                                table1.SetWidths(firstTablecellwidth);
                                                doc.Add(table1);
                                                strCrntDt = strCreatedDate;
                                            }
                                            MsgCreatedUserName = string.Empty;
                                            if (msgResult.user != null)
                                            {
                                                MsgCreatedUserName = msgResult.user.ToString();
                                            }
                                            //Get created by User dispaly form slack
                                            MsgCreatedByUserName = GetUserDisplayName(JsonUserList, MsgCreatedUserName);
                                             UserMessage = string.Empty;
                                            if (msgResult.text != null)
                                            {
                                                UserMessage = msgResult.text.ToString();
                                                if (msgResult.attachments.Count > 0)
                                                {
                                                    UserMessage = UserMessage + "\n" + msgResult.attachments.FirstOrDefault().text; 
                                                }
                                            }
                                            if ((msgResult.subtype != null) && (msgResult.subtype.ToString() == "file_share") && (arrSubType.Contains(msgResult.subtype.ToString())))
                                            {
                                                Regex regex = new Regex(@"\<.*?\>");
                                                MatchCollection matches = regex.Matches(UserMessage);
                                                string fls = msgResult.SlackFiles.FirstOrDefault().url_private_download;
                                                string strOldUrl = matches[1].ToString();
                                                string strNewVal = "<" + fls + ">";
                                                UserMessage = UserMessage.Replace(strOldUrl, strNewVal);
                                            }
                                            if ((msgResult.subtype != null) && (msgResult.subtype.ToString() == "pinned_item") && (arrSubType.Contains(msgResult.subtype.ToString())))
                                            {

                                                UserMessage = msgResult.pinned_item.FirstOrDefault().comment;
                                            }
                                            Paragraph txtUserName = new Paragraph(MsgCreatedByUserName, userNameFont);
                                            Console.WriteLine("User name:" + MsgCreatedByUserName + " " + strCreatedTime, userNameFont);
                                            Chunk cCreatedTime = new Chunk(" " + strCreatedTime, georgiaTimeFont);
                                            txtUserName.Add(cCreatedTime);
                                            doc.Add(txtUserName);
                                            Console.WriteLine("User Message:" + UserMessage);
                                            Font georgia = FontFactory.GetFont("Arial", 7f);
                                            georgia.Color = BaseColor.BLACK;
                                            doc.Add(new Paragraph(UserMessage, georgia));
                                            Console.WriteLine("Created On:" + UnixTimeStampToDateTime(Convert.ToDouble(UserCreatedDtTime)));
                                            doc.Add(new Paragraph("\n"));
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        strBuilderMain.AppendLine(ex.Message);
                                        Logger.LogException("Error while processing Channel Name:" + Channel_name + ":" + ex.Message.ToString());
                                        MailUtility objmail = new MailUtility();
                                        objmail.SendEmail("Error while processing Channel Name:" + Channel_name, ex.Message.ToString());
                                    }
                                    //}//Messages Innver loop
                               
                                }//Messages loop
                               
                                doc.Close();
                                if (File.Exists(System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.System), SlackArchivePath + "TempChannel" + ".pdf")))
                                {
                                     designationFile = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.System), SlackArchivePath + Channel_name + ".pdf");
                                     souceFile = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.System), SlackArchivePath + "TempChannel" + ".pdf");
                                     souceFile2 = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.System), SlackArchivePath + "TempChannel2" + ".pdf");
                                    // MergeChannelPdfFiles(souceFile,designationFile);
                                    if (File.Exists(designationFile))
                                    {
                                        File.Copy(designationFile, souceFile2, true);
                                        File.Delete(designationFile);
                                    }
                                    List<string> source = new List<string>();
                                    if (File.Exists(souceFile))
                                    {
                                        source.Add(souceFile);
                                    }
                                    if (File.Exists(souceFile2))
                                    {
                                        source.Add(souceFile2);
                                    }
                                  
                                    //MergeFilesFinal(designationFile, source);
                                    MergeSlackPdfFiles(source, designationFile);

                                }

                            }

                        }
                    }
                    else
                    {
                        //
                        Console.WriteLine("No Messages to Process");
                    }

                }//channel Loop
                File.SetLastWriteTime(LastRuntime, DateTime.Now);
                //Console.ReadLine();

            }
            catch (Exception ex)
            {
                strBuilderMain.AppendLine(ex.Message);
            }
            finally
            {
                strBuilderMain.AppendLine("Process execution completed:" + DateTime.Now);
                Logger.LogInformation(strBuilderMain.ToString());
                strBuilderMain.Clear();
            }

        }

        /// <summary>
        /// Merging the slack messges PDF with Old channel messges
        /// </summary>
        /// <param name="files"></param>
        /// <param name="output"></param>
        public static void MergeSlackPdfFiles(IEnumerable<string> files, string output)
        {
            iTextSharp.text.Document doc;
            iTextSharp.text.pdf.PdfCopy pdfCpy;

            doc = new iTextSharp.text.Document();
            pdfCpy = new iTextSharp.text.pdf.PdfCopy(doc, new System.IO.FileStream(output, System.IO.FileMode.Create));
            doc.Open();
            foreach (string file in files)
            {
                // initialize a reader
                iTextSharp.text.pdf.PdfReader reader = new iTextSharp.text.pdf.PdfReader(file);
                int pageCount = reader.NumberOfPages;
                // set page size for the documents
                doc.SetPageSize(reader.GetPageSizeWithRotation(1));
                for (int pageNum = 1; pageNum <= pageCount; pageNum++)
                {
                    iTextSharp.text.pdf.PdfImportedPage page = pdfCpy.GetImportedPage(reader, pageNum);
                    pdfCpy.AddPage(page);
                }

                reader.Close();
            }

            doc.Close();
            foreach (string file in files)
            {
                File.Delete(file);
            }

        }


   /// <summary>
       /// Setting hearder wiht for line separater table
       /// </summary>
       /// <param name="font"></param>
       /// <param name="headers"></param>
       /// <returns></returns>
    public static float[] GetHeaderWidths(Font font, params string[] headers)
        {
            var total = 0;
            var columns = headers.Length;
            var widths = new int[columns];
            for (var iheadLenght = 0; iheadLenght < columns; ++iheadLenght)
            {
                var w = font.GetCalculatedBaseFont(true).GetWidth(headers[iheadLenght]);
                total += w;
                widths[iheadLenght] = w;
            }
            var result = new float[columns];
            for (var iWidth = 0; iWidth < columns; ++iWidth)
            {
                result[iWidth] = (float)widths[iWidth] / total * 100;
            }
            return result;
        }

        
        /// <summary>
        /// converting to datetime format
        /// </summary>
        /// <param name="unixTimeStamp"></param>
        /// <returns></returns>
        public static DateTime UnixTimeStampToDateTime(double unixTimeStamp)
        {
            try
            {
                // Unix timestamp is seconds past epoch
                System.DateTime dtDateTime = new DateTime(1970, 1, 1, 0, 0, 0, 0, System.DateTimeKind.Utc);
                dtDateTime = dtDateTime.AddSeconds(unixTimeStamp).ToLocalTime();
                return dtDateTime;

            }
            catch (Exception ex)
            {

                throw;
            }


        }


        /// <summary>
        /// Converting to datetime
        /// </summary>
        /// <param name="StrDateTime"></param>
        /// <returns></returns>
        public static double Datetimeconversion(string StrDateTime)
        {
            // string dateString = "2/23/2018 12:08:00 PM";// "2/22/2018 10:30:00 PM";
            double unixconvertsion = 0;
            DateTime dateTime =
                DateTime.Parse(StrDateTime, System.Globalization.CultureInfo.InvariantCulture);
            //Console.WriteLine(dateTime.ToString());

            // DateTime dateTime = new DateTime();

            DateTime unixStart = new DateTime(1970, 1, 1, 0, 0, 0, 0, System.DateTimeKind.Utc);
            long unixTimeStampInTicks = (dateTime.ToUniversalTime() - unixStart).Ticks;
            unixconvertsion = (double)unixTimeStampInTicks / TimeSpan.TicksPerSecond;
            return unixconvertsion;

        }
        

        /// <summary>
        /// Parse Slack display Name
        /// </summary>
        /// <param name="unixTimeStamp"></param>
        /// <returns></returns>
        public static string GetUserDisplayName(string JsonResopnse, string slackUserId)
        {
            string SlackUserDisplayName = "";
            try
            {
                JObject o = JObject.Parse(JsonResopnse);
                JArray arr = (JArray)o.SelectToken("members");
                var name = arr.FirstOrDefault(x => x.Value<string>("id") == "" + slackUserId + "").Value<string>("real_name");
                return SlackUserDisplayName = name;

            }
            catch (Exception ex)
            {

                throw;
            }


        }


    }
   




}
