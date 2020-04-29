using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json;
using System.Data.SQLite;
using RestSharp;
using System.IO;
using System.Windows.Forms;
using agi = HtmlAgilityPack;

namespace keywordGOGO
{
    class InstagramAPI
    {

        public List<InstagramTagWordList> InstagramJsonDataSet(string tag)
        {
            string url = "https://www.instagram.com/explore/tags/" + tag + "/?__a=1";

            List<string> InstaTagList = new List<string>();
            List<InstagramTagWordList> InstaTagWordList = new List<InstagramTagWordList>();
            string jsonDataset = HttpWebRequestText(url);

            if (jsonDataset != string.Empty)
            {
                JObject obj = JObject.Parse(jsonDataset);
                JObject graphqlObj = JObject.Parse(obj["graphql"].ToString());
                JObject hashtagObj = JObject.Parse(graphqlObj["hashtag"].ToString());

                InstaTagList.AddRange(edgeHashtagToMedia(hashtagObj));
                InstaTagList.AddRange(edgeHashtagToTopPosts(hashtagObj));

                // 중복 단어의 수를 체크한다.
                var q = InstaTagList.GroupBy(x => x)
               .Select(g => new { Value = g.Key, Count = g.Count() })
               .OrderByDescending(x => x.Count).ToList();
                //중복 키워드를 리스트에 담는다.
                foreach (var temp in q)
                {
                    InstaTagWordList.Add(new InstagramTagWordList() { value = temp.Value, count = temp.Count });
                }
            }

            return InstaTagWordList;
        }

        public List<string> edgeHashtagToTopPosts(JObject hashtagObj)
        {
            List<string> InstaTagList = new List<string>();
            JObject edgesObj = JObject.Parse(hashtagObj["edge_hashtag_to_top_posts"].ToString());
            JArray array = JArray.Parse(edgesObj["edges"].ToString());
            foreach (JObject edgesdata in array)
            {
                //태그추출
                JObject indexData = JObject.Parse(edgesdata["node"].ToString());
                JObject captionData = JObject.Parse(indexData["edge_media_to_caption"].ToString());
                JArray tagarray = JArray.Parse(captionData["edges"].ToString());
                foreach (JObject TagObject in tagarray)
                {
                    //Console.WriteLine(TagObject["node"]["text"]);
                    string tagdata = TagObject["node"]["text"].ToString();
                    int sbindex = tagdata.IndexOf("#");

                    if (sbindex != -1)
                    {
                        StringBuilder sb = new StringBuilder(tagdata);

                        string resultTagAll = sb.ToString(sbindex, tagdata.Length - sbindex);

                        string[] result = resultTagAll.Split(new string[] { "#" }, StringSplitOptions.RemoveEmptyEntries);
                        for (int i = 0; i < result.Length; i++)
                        {
                            int lastsbindex = result[i].ToString().IndexOf("\n");

                            StringBuilder stb = new StringBuilder(result[i]);
                            if (lastsbindex == -1)
                            {
                                StringBuilder blank = new StringBuilder(result[i]);
                                int blankindex = blank.ToString().IndexOf(" ");
                                if (blankindex == -1)
                                {
                                    string tagresult = blank.ToString();
                                    InstaTagList.Add(tagresult);
                                    Console.WriteLine("++++++++++++++++++++++++++");
                                    Console.WriteLine(sb);
                                    Console.WriteLine("--------------------------");
                                    Console.WriteLine(tagresult);
                                }
                                else
                                {
                                    string tagresult = stb.ToString(0, blankindex);
                                    InstaTagList.Add(tagresult.Replace("\n", ""));
                                    Console.WriteLine("++++++++++++++++++++++++++");
                                    Console.WriteLine(sb);
                                    Console.WriteLine("--------------------------");
                                    Console.WriteLine(tagresult.Replace("\n", ""));

                                }
                            }
                            else
                            {
                                string tagresult = stb.ToString(0, lastsbindex);
                                InstaTagList.Add(tagresult);
                                Console.WriteLine("++++++++++++++++++++++++++");
                                Console.WriteLine(sb);
                                Console.WriteLine("--------------------------");
                                Console.WriteLine(tagresult);
                            }

                        }
                    }
                }
            }

            return InstaTagList;

        }

        public List<string> edgeHashtagToMedia(JObject hashtagObj)
        {
            List<string> InstaTagList = new List<string>();
            JObject edgesObj = JObject.Parse(hashtagObj["edge_hashtag_to_media"].ToString());
            JArray array = JArray.Parse(edgesObj["edges"].ToString());
            foreach (JObject edgesdata in array)
            {
                //태그추출
                JObject indexData = JObject.Parse(edgesdata["node"].ToString());
                JObject captionData = JObject.Parse(indexData["edge_media_to_caption"].ToString());
                JArray tagarray = JArray.Parse(captionData["edges"].ToString());
                foreach (JObject TagObject in tagarray)
                {
                    //Console.WriteLine(TagObject["node"]["text"]);
                    string tagdata = TagObject["node"]["text"].ToString();
                    int sbindex = tagdata.IndexOf("#");

                    if (sbindex != -1)
                    {
                        StringBuilder sb = new StringBuilder(tagdata);

                        string resultTagAll = sb.ToString(sbindex, tagdata.Length - sbindex);

                        string[] result = resultTagAll.Split(new string[] { "#" }, StringSplitOptions.RemoveEmptyEntries);
                        for (int i = 0; i < result.Length; i++)
                        {
                            int lastsbindex = result[i].ToString().IndexOf("\n");

                            StringBuilder stb = new StringBuilder(result[i]);
                            if (lastsbindex == -1)
                            {
                                StringBuilder blank = new StringBuilder(result[i]);
                                int blankindex = blank.ToString().IndexOf(" ");
                                if (blankindex == -1)
                                {
                                    string tagresult = blank.ToString();
                                    InstaTagList.Add(tagresult);
                                    Console.WriteLine("++++++++++++++++++++++++++");
                                    Console.WriteLine(sb);
                                    Console.WriteLine("--------------------------");
                                    Console.WriteLine(tagresult);
                                }
                                else
                                {
                                    string tagresult = stb.ToString(0, blankindex);
                                    InstaTagList.Add(tagresult.Replace("\n", ""));
                                    Console.WriteLine("++++++++++++++++++++++++++");
                                    Console.WriteLine(sb);
                                    Console.WriteLine("--------------------------");
                                    Console.WriteLine(tagresult.Replace("\n", ""));
                                }
                            }
                            else
                            {
                                string tagresult = stb.ToString(0, lastsbindex);
                                InstaTagList.Add(tagresult);
                                Console.WriteLine("++++++++++++++++++++++++++");
                                Console.WriteLine(sb);
                                Console.WriteLine("--------------------------");
                                Console.WriteLine(tagresult);
                            }

                        }
                    }
                }
            }
            return InstaTagList;
        }


            /// <summary>
            /// 타겟 URL 부터 HTML 코드를 가져온다.
            /// </summary>
            /// <param name="tagetUrl">타겟 URL</param>
            /// <returns>String으로 된 HTML 소스 </returns>
            public string HttpWebRequestText(string tagetUrl)
        {
            string responseText = string.Empty;
            string url = tagetUrl;
            Thread.Sleep(1000);
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);
            request.Method = "GET";
            request.Timeout = 30 * 1000; // 30초
            request.Headers.Add("Authorization", "BASIC SGVsbG8="); // 헤더 추가 방법
            try
            {
                using (HttpWebResponse resp = (HttpWebResponse)request.GetResponse())
                {
                    HttpStatusCode status = resp.StatusCode;

                    Console.WriteLine(status);  // 정상이면 "OK"

                    //listBox1.Items.Add("네이버와 통신결과 : "+status);

                    Stream respStream = resp.GetResponseStream();
                    using (StreamReader sr = new StreamReader(respStream))
                    {
                        responseText = sr.ReadToEnd();
                    }
                }
            }catch
            {
                responseText = string.Empty;
            }
            return responseText;
        }
    }
}
