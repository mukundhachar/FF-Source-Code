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
using System.Configuration;
using System.IO;
using iTextSharp;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.Text.RegularExpressions;


namespace SlackApiAccess
{
    class Program
    {
        static void Main(string[] args)
        {
            //string outputFile = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "test.pdf");
         
            StringBuilder strBuilderMain = new StringBuilder();
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
            //Creating last runtime file if not exist
            if (!File.Exists(LastRuntime))
            {
                File.Create(LastRuntime);
            }
            DateTime? Oldest = null;
            if (Utility.GetConfigValueByKey("Olddest") == "null")
            {
                Oldest = null;
            }
            else
            {
                Oldest = Convert.ToDateTime(Utility.GetConfigValueByKey("Olddest"));

               // Oldest = null;
                //getting the last run time
                DateTime dt = File.GetLastWriteTime(LastRuntime);


                Oldest = dt;
            }
            double lngOldest = 0;
            if (Oldest == null)
            {
                lngOldest = 0;

            }
            else
            {
               // lngOldest = ConvertToUnixTime(Convert.ToDateTime(Oldest));
                lngOldest = Datetimeconversion(Convert.ToString(Oldest));
                
                //lngOldest = ConvertToUnixTime(Convert.ToDateTime("2/22/2018 6:00:00 PM"));
            }          

            try
            {
                SlackClient client = new SlackClient(LegacyToken);
                //Get channel information form slack
                string JsonChannelsStr = client.GetChannelInfoFromSlack();
                //*******************Filtring the Channel List********************************************
                JObject objChannels = JObject.Parse(JsonChannelsStr);
                var listJChannelobject = new List<JObject>();
                string[] arrChannels = FilterChannels == string.Empty ? null : FilterChannels.Split(new string[] { "," }, StringSplitOptions.None);
                if (arrChannels != null)
                {
                    foreach (var s in arrChannels)
                    {
                        var temp = (objChannels["channels"].Values<JObject>()
                                 .Where(m => m["name"].Value<string>().Contains(s)))
                                 .ToList();
                        temp.ForEach(a => listJChannelobject.Add(a));
                    }
                }
                else
                {
                    var temp = (objChannels["channels"].Values<JObject>()
                              ).ToList();
                    temp.ForEach(a => listJChannelobject.Add(a));
                }
                //***********************************************************************************
                Thread.Sleep(TimeSpan.FromSeconds(5));
                //Getting the user list form Slack
                string JsonUserList = client.GetSlackUserList();
                JObject results = JObject.Parse(JsonChannelsStr);
                // Loop through each channel
                ////foreach (var result in results["channels"])
                foreach (JObject result in listJChannelobject)
                {
                    string Channel_name = (string)result["name_normalized"];
                    string channelId = (string)result["id"];
                    //outputFile=Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), Channel_name+".pdf");
                    Console.WriteLine("Channel name:" + Channel_name);
                    //////Font titleFont = FontFactory.GetFont(FontFactory.TIMES_ROMAN, 12, Font.BOLD);
                    //Font titleFont = FontFactory.GetFont(FontFactory.a, 16, Font.BOLD );
                    Font titleFont = new Font(FontFactory.GetFont("Microsoft Sans Serif", 10, Font.BOLD));
                    titleFont.SetColor(44, 45, 48);
                    Font regularFont = FontFactory.GetFont("Microsoft Sans Serif", 8);
                    ///////Font userNameFont = FontFactory.GetFont(FontFactory.TIMES_ROMAN, 10, Font.BOLD);
                    Font userNameFont = new Font(FontFactory.GetFont("Microsoft Sans Serif", 8, Font.BOLD));
                    userNameFont.SetColor(44, 45, 48);
                    Paragraph title;
                    title = new Paragraph("#" + Channel_name, titleFont);
                    title.Alignment = Element.ALIGN_LEFT;
                   
                    Console.WriteLine("                                                                  ");
                    Console.WriteLine("                                                                  ");
                    Thread.Sleep(TimeSpan.FromSeconds(5));
                    //Get all channel messages for the given channel ID
                    string channelMessages = client.GetChannelMessagesInfoFromSlack(channelId, lngOldest);
                    //string channelMessages = client.GetChannelMessagesInfoFromSlack(channelId, 1519344000);
                    JObject objMsgResults = JObject.Parse(channelMessages);
                    var listJobject = new List<JObject>();
                    string[] arrEmojis = FilterEmoji == string.Empty ? null : FilterEmoji.Split(new string[] { "," }, StringSplitOptions.None);
                    if (arrEmojis != null)
                    {


                        //var temp2 = (objMsgResults["messages"].Values<JObject>()
                        //            .Where(m =>Convert.ToDateTime(UnixTimeStampToDateTime(Convert.ToDouble(m["edited"].Children()["ts"].ToString()))) > Oldest
                        //                //m["text"].Value<string>().Contains(s)
                        //               ////&& Convert.ToDateTime(UnixTimeStampToDateTime(Convert.ToDouble(m["ts"].ToString()))) > Oldest
                        //              // &&

                        //                    )
                        //            )
                        //            .ToList();


                        foreach (var s in arrEmojis)
                        {
                            var temp = (objMsgResults["messages"].Values<JObject>()
                                     .Where(m => m["text"].Value<string>().Contains(s)
                                ////&& Convert.ToDateTime(UnixTimeStampToDateTime(Convert.ToDouble(m["ts"].ToString()))) > Oldest
                                //&& Convert.ToDateTime(UnixTimeStampToDateTime(Convert.ToDouble(m["edited"]["ts"].ToString()))) > Oldest

                                         )
                                     )
                                     .ToList();
                            temp.ForEach(a => listJobject.Add(a));
                        }
                    }
                    else
                    {
                        var temp = (objMsgResults["messages"].Values<JObject>()).ToList();
                        temp.ForEach(a => listJobject.Add(a));

                        //var temp = (objMsgResults["messages"].Values<JObject>()).ToList();
                        //     //.Where(m => m["is_starred"].Value<string>() == "true")).ToList();
                        //temp.ForEach(a => listJobject.Add(a));

                    }






                    if (listJobject.Count > 0)
                    {

                        outputFile = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.System), SlackArchivePath + Channel_name + ".pdf");
                        using (FileStream fs = new FileStream(outputFile, File.Exists(outputFile) ? FileMode.Append : FileMode.Create, FileAccess.Write, FileShare.None))
                        {
                            Document doc = new Document(PageSize.LETTER);
                            using (PdfWriter writer = PdfWriter.GetInstance(doc, fs))
                            {
                                doc.Open();
                                doc.Add(title);
                                // Loop through each message
                                string strCrntDt = "";
                                int PadLeft = -38;
                                //foreach (var msgResult in objMsgResults["messages"])
                                foreach (JObject msgResult in listJobject)
                                {
                                    PadLeft = -140;
                                    //foreach (var msgResult in ObjMsgResult["messages"])
                                    //{
                                    try
                                    {
                                        string MsgCreatedByUserName = string.Empty;
                                        if (msgResult["user"] != null)
                                        {
                                            string UserCreatedDtTime = string.Empty;
                                            if (msgResult["ts"] != null)
                                            {
                                                UserCreatedDtTime = msgResult["ts"].ToString();
                                            }
                                            string strCreatedTime = Convert.ToDateTime(UnixTimeStampToDateTime(Convert.ToDouble(UserCreatedDtTime))).ToString("h:mm tt");
                                            string strCreatedDate = Convert.ToDateTime(UnixTimeStampToDateTime(Convert.ToDouble(UserCreatedDtTime))).ToString("dddd MMMM dd");
                                            Font georgiaTimeFont = FontFactory.GetFont("Arial", 7f);
                                            georgiaTimeFont.Color = BaseColor.GRAY;
                                            if (strCrntDt != strCreatedDate)
                                            {
                                                //************************************************************************************************************
                                                string[] headers = new string[] { " ", " ", " " };
                                                //Font font = new Font("Arial", 8, Font.BOLD);
                                                Font font = new Font(FontFactory.GetFont("Microsoft Sans Serif", 8, Font.BOLD));
                                                font.SetColor(44, 45, 48);
                                                var table1 = new PdfPTable(headers.Length) { WidthPercentage = 100 };
                                                table1.SetWidths(GetHeaderWidths(font, headers));
                                                Paragraph cell3Text1;
                                                for (int i = 0; i < headers.Length; ++i)
                                                {

                                                    // table.AddCell(new PdfPCell(new Phrase( headers[i], font)));
                                                    if (i == 1)
                                                    {
                                                        PdfPCell cell = new PdfPCell(new Phrase(strCreatedDate, font));
                                                        cell.Border = 0;
                                                        //cell.BorderWidthBottom = 3f;
                                                        //cell.BorderWidthTop = 3f;
                                                        cell.PaddingBottom = 10f;
                                                        //cell.PaddingLeft = 20f;
                                                        cell.PaddingRight = 0;
                                                        cell.PaddingLeft = 0;
                                                        cell.PaddingTop = 4f;
                                                        cell.HorizontalAlignment = (Element.ALIGN_CENTER);
                                                        //cell.Width = strCreatedDate.Length;
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
                                                //**************************************************************************************************************
                                                strCrntDt = strCreatedDate;
                                                //doc.Add(pCreatedDate);
                                                //Paragraph p = new Paragraph(new Chunk(new iTextSharp.text.pdf.draw.LineSeparator(0.0F, 100.0F, BaseColor.LIGHT_GRAY, Element.ALIGN_LEFT, 1)));
                                                //doc.Add(p);
                                            }
                                            string MsgCreatedUserName = string.Empty;
                                            if (msgResult["user"] != null)
                                            {
                                                MsgCreatedUserName = msgResult["user"].ToString();
                                            }
                                            //Get created by User dispaly form slack
                                            //Thread.Sleep(TimeSpan.FromSeconds(5));
                                            //////////////// MsgCreatedByUserName = client.GetUserDisplayNameFromSlack(MsgCreatedUserName);
                                            MsgCreatedByUserName = GetUserDisplayName(JsonUserList, MsgCreatedUserName);
                                            //Thread.Sleep(TimeSpan.FromSeconds(2));
                                            string UserMessage = string.Empty;
                                            if (msgResult["text"] != null)
                                            {
                                                UserMessage = msgResult["text"].ToString();
                                            }
                                            if ((msgResult["subtype"] != null) && (msgResult["subtype"].ToString() == "file_share"))
                                            {
                                                Regex regex = new Regex(@"\<.*?\>");
                                                MatchCollection matches = regex.Matches(UserMessage);
                                                string fls = msgResult["file"].ToString();
                                                dynamic flarray = JObject.Parse(fls);
                                                string urlPrivateDownload = flarray.url_private_download.ToString();
                                                string strOldUrl = matches[1].ToString();
                                                string strNewVal = "<" + urlPrivateDownload + ">";
                                                UserMessage = UserMessage.Replace(strOldUrl, strNewVal);
                                            }
                                            Paragraph txtUserName = new Paragraph(MsgCreatedByUserName, userNameFont);
                                            Console.WriteLine("User name:" + MsgCreatedByUserName + " " + strCreatedTime, userNameFont);
                                            Chunk cCreatedTime = new Chunk(" " + strCreatedTime, georgiaTimeFont);
                                            txtUserName.Add(cCreatedTime);
                                            doc.Add(txtUserName);
                                            Console.WriteLine("User Message:" + UserMessage);
                                            Font georgia = FontFactory.GetFont("Arial", 7f);
                                            georgia.Color = BaseColor.GRAY;
                                            doc.Add(new Paragraph(UserMessage, georgia));
                                            Console.WriteLine("Created On:" + UnixTimeStampToDateTime(Convert.ToDouble(UserCreatedDtTime)));
                                            // doc.Add(new Paragraph("Created On:" + UnixTimeStampToDateTime(Convert.ToDouble(UserCreatedDtTime))));
                                            // Console.WriteLine("                                                                  ");
                                            doc.Add(new Paragraph("\n"));
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        strBuilderMain.AppendLine(ex.Message);

                                        Logger.LogException("Error while processing Channel Name:" + Channel_name + ":" + ex.Message.ToString());
                                        try
                                        {
                                            MailUtility objmail = new MailUtility();
                                            objmail.SendEmail("Error while processing Channel Name:" + Channel_name, ex.Message.ToString());
                                        }
                                        catch (Exception exm)
                                        {
                                            Logger.LogException("Channel name:" + Channel_name + ":" + exm.Message.ToString());
                                        }


                                    }
                                    //}//Messages Innver loop
                                }
                                doc.Close();
                            }

                        }
                    }
                    else
                    {
                        Console.WriteLine("**************************No Messages to Process***********************");
                    }
                  
                    Console.WriteLine("**************************************************************************");

                }//channel Loop

                //Setting the last runtime
               File.SetLastWriteTime(LastRuntime, DateTime.Now);
               
               
                //******************************************************************
                Console.ReadLine();

            }
            catch (Exception ex)
            {
                //doc.Close();

                strBuilderMain.AppendLine(ex.Message);
                //The remote server returned an error: (429) Too Many Requests.
            }
            finally
            {
                strBuilderMain.AppendLine("Process execution completed:" + DateTime.Now);
                strBuilderMain.AppendLine("----------------------------------------------------------");
                Logger.LogInformation(strBuilderMain.ToString());
                strBuilderMain.Clear();
            }




        }




        public static float[] GetHeaderWidths(Font font, params string[] headers)
        {
            var total = 0;
            var columns = headers.Length;
            var widths = new int[columns];
            for (var i = 0; i < columns; ++i)
            {
                var w = font.GetCalculatedBaseFont(true).GetWidth(headers[i]);
                total += w;
                widths[i] = w;
            }
            var result = new float[columns];
            for (var i = 0; i < columns; ++i)
            {
                result[i] = (float)widths[i] / total * 100;
            }
            return result;
        }

        public static void DrawDashedLines(Document doc, PdfWriter writer)
        {
            PdfContentByte cb = writer.DirectContent;
            cb.SetLineDash(3f, 3f);
            cb.MoveTo(0, doc.PageSize.Height / 3);
            cb.LineTo(doc.PageSize.Width, doc.PageSize.Height / 3);
            cb.Stroke();
            // some code removed for clarity
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

                throw ex;
            }


        }



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

                throw ex;
            }


        }




        private static string MakeLink(string txt)
        {
            Regex regx = new Regex("http://([\\w+?\\.\\w+])+([a-zA-Z0-9\\~\\!\\@\\#\\$\\%\\^\\&amp;\\*\\(\\)_\\-\\=\\+\\\\\\/\\?\\.\\:\\;\\'\\,]*)?", RegexOptions.IgnoreCase);
            MatchCollection mactches = regx.Matches(txt);
            foreach (Match match in mactches)
            {
                txt = txt.Replace(match.Value, "<a href='" + match.Value + "'>" + match.Value + "</a>");
            }
            return txt;
        }


        /// <summary>
        /// Convert a date time object to Unix time representation.
        /// </summary>
        /// <param name="datetime">The datetime object to convert to Unix time stamp.</param>
        /// <returns>Returns a numerical representation (Unix time) of the DateTime object.</returns>
        public static long ConvertToUnixTime(DateTime datetime)
        {
            DateTime sTime = new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc);

            return (long)(datetime - sTime).TotalSeconds;
        }


    }


}
