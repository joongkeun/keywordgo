using Newtonsoft.Json.Linq;
using RestSharp;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using agi = HtmlAgilityPack;

namespace keywordGOGO
{
    class NaverApi
    {
        // 메인폼 전달 이벤트 선언
        public static event listBoxText ReturnToMessage;
        public static event labelText ReturnToLabel;
        iniUtil ini = new iniUtil(Application.StartupPath + "\\config.ini");


        /// <summary>
        /// 네이버 광고 api 접속하여 JSON 형식으로 자료를 받아온다.
        /// </summary>
        /// <param name="query"></param>
        /// <returns></returns>
        public string NaverAdApi(string query)
        {
            try
            {
                var baseUrl = "https://api.naver.com";
                var apiKey = ini.GetIniValue("ADAPI", "apiKey");
                var secretKey = ini.GetIniValue("ADAPI", "secretKey");
                var managerCustomerId = long.Parse(ini.GetIniValue("ADAPI", "managerCustomerId"));

                var rest = new SearchAdApi(baseUrl, apiKey, secretKey);

                var request = new RestRequest("/keywordstool", Method.GET);
                request.AddQueryParameter("hintKeywords", query);
                request.AddQueryParameter("showDetail", "1");
                string jSonData = rest.Execute<List<RelKwdStat>>(request, managerCustomerId);
                return jSonData;
            }
            catch
            {
                ReturnToMessage("-------------------------------------------");
                ReturnToMessage("네이버광고에서 데이터를 불러오지 못했습니다.");
                ReturnToMessage("광고 API 정보를 다시확인해주세요.");
                ReturnToMessage("-------------------------------------------");
                string jSonData = "";
                return jSonData;
            }
        }


        /// <summary>
        /// 네이버 광고 api 접속하여 JSON 형식으로 자료를 받아온다.
        /// </summary>
        /// <returns></returns>
        public string NaverAdBizmoneyApi()
        {
            try
            {
                var baseUrl = "https://api.naver.com";
                var apiKey = ini.GetIniValue("ADAPI", "apiKey");
                var secretKey = ini.GetIniValue("ADAPI", "secretKey");
                var managerCustomerId = long.Parse(ini.GetIniValue("ADAPI", "managerCustomerId"));

                var rest = new SearchAdApi(baseUrl, apiKey, secretKey);

                var request = new RestRequest("/billing/bizmoney", Method.GET);
                //request.AddQueryParameter("hintKeywords", query);
                //request.AddQueryParameter("showDetail", "1");
                string jSonData = rest.Execute<Bizmoney>(request, managerCustomerId);
           
                return jSonData;
            }
            catch
            {
                ReturnToMessage("-------------------------------------------");
                ReturnToMessage("네이버광고에서 데이터를 불러오지 못했습니다.");
                ReturnToMessage("광고 API 정보를 다시확인해주세요.");
                ReturnToMessage("-------------------------------------------");
                string jSonData = "";
                return jSonData;
            }
        }


        /// <summary>
        /// 네이버에서 자료를 조회하여 JSON 형식으로 자료를 받아온다.
        /// </summary>
        public string NaverOpenApi(string query)
        {
            
            Thread.Sleep(300);
            string url = "https://openapi.naver.com/v1/search/shop.json?query=" + query + "&display=80&sort=sim"; // 결과가 JSON 포맷
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);
            request.Headers.Add("X-Naver-Client-Id", ini.GetIniValue("OPENAPI", "ClientId")); // 클라이언트 아이디
            request.Headers.Add("X-Naver-Client-Secret", ini.GetIniValue("OPENAPI", "ClientSecret"));       // 클라이언트 시크릿
            HttpWebResponse response = (HttpWebResponse)request.GetResponse();
            string status = response.StatusCode.ToString();
            if (status == "OK")
            {
                Stream stream = response.GetResponseStream();
                StreamReader reader = new StreamReader(stream, Encoding.UTF8);
                string text = reader.ReadToEnd();
                // Console.WriteLine(text);
                return text;
            }
            else
            {
                return null;
            }
        }

        /// <summary>
        /// 오픈api를 접속하여 받아 판매상품수와 샵정보를 리턴한다.
        /// </summary>
        /// <param name="keyWord"></param>
        /// <returns></returns>
        public OpenApiDataSetResult OpenApiDataSet(string keyWord)
        {
            int total = 0; // 총검색량
            List<string> titleKeywordListResult = new List<string>();
            OpenApiDataSetResult Result = new OpenApiDataSetResult();
            List<ShopAPIResult> ShopResult = new List<ShopAPIResult>();
        
            string naver = NaverOpenApi(keyWord);
            JObject obj = JObject.Parse(naver);
            total = Convert.ToInt32(obj["total"]);

            JArray array = JArray.Parse(obj["items"].ToString());

            if (total == 0) //같은 물품이 없는 경우
            {
                string title = "";//검색 결과 문서의 제목을 나타낸다. 제목에서 검색어와 일치하는 부분은 태그로 감싸져 있다.
                string link = "";//검색 결과 문서의 하이퍼텍스트 link를 나타낸다.
                string image = "";//썸네일 이미지의 URL이다. 이미지가 있는 경우만 나타난다.
                string lprice = "";//최저가 정보이다. 
                string hprice = "";//최고가 정보이다. 최고가 정보가 없거나 가격비교 데이터가 없는 경우 0으로 표시된다.
                string mallName = "";//상품을 판매하는 쇼핑몰의 상호이다. 정보가 없을 경우 네이버로 표기된다.
                string productId = "";//해당 상품에 대한 ID 이다.
                string productType = "";//상품군 정보를 일반상품, 중고상품, 단종상품, 판매예정상품으로 구분한다.

                ShopResult.Add(new ShopAPIResult
                {
                    RelKeyword = keyWord,
                    Title = title,
                    Image = image,
                    Lprice = lprice,
                    Hprice = hprice,
                    MallName = mallName,
                    ProductId = productId,
                    ProductType = productType

                });

            }
            else // 같은 물품이 다수가 있는 경우 
            {
                int roop = 0;
                foreach (JObject itemObj in array)
                {
                    if (total == 1) // 검색된 데이터가 한건
                    {

                        string title = itemObj["title"].ToString();//검색 결과 문서의 제목을 나타낸다. 제목에서 검색어와 일치하는 부분은 태그로 감싸져 있다.
                        string link = itemObj["link"].ToString();//검색 결과 문서의 하이퍼텍스트 link를 나타낸다.
                        string image = itemObj["image"].ToString();//썸네일 이미지의 URL이다. 이미지가 있는 경우만 나타난다.
                        string lprice = itemObj["lprice"].ToString();//최저가 정보이다. 
                        string hprice = itemObj["hprice"].ToString();//최고가 정보이다. 최고가 정보가 없거나 가격비교 데이터가 없는 경우 0으로 표시된다.
                        string mallName = itemObj["mallName"].ToString();//상품을 판매하는 쇼핑몰의 상호이다. 정보가 없을 경우 네이버로 표기된다.
                        string productId = itemObj["productId"].ToString();//해당 상품에 대한 ID 이다.
                        string productType = itemObj["productType"].ToString();//상품군 정보를 일반상품, 중고상품, 단종상품, 판매예정상품으로 구분한다.

                        if (lprice.Equals(""))
                        {
                            lprice = "0";
                        }

                        if (hprice.Equals(""))
                        {
                            hprice = "0";
                        }


                        title = title.Replace("<b>", "").Replace("</b>", "");
                        List<string> titleKeywordList = title.Split(' ').ToList();
                        titleKeywordListResult.AddRange(titleKeywordList);
                        ShopResult.Add(new ShopAPIResult
                        {
                            RelKeyword = keyWord,
                            Title = title,
                            Image = image,
                            Lprice = lprice,
                            Hprice = hprice,
                            MallName = mallName,
                            ProductId = productId,
                            ProductType = productType,
                            TitleKeywordList = titleKeywordList

                        });
                    }
                    else // 검색된 데이터가 다수 인 경우
                    {

                        string title = itemObj["title"].ToString();//검색 결과 문서의 제목을 나타낸다. 제목에서 검색어와 일치하는 부분은 태그로 감싸져 있다.
                        string link = itemObj["link"].ToString();//검색 결과 문서의 하이퍼텍스트 link를 나타낸다.
                        string image = itemObj["image"].ToString();//썸네일 이미지의 URL이다. 이미지가 있는 경우만 나타난다.
                        string lprice = itemObj["lprice"].ToString();//최저가 정보이다. 
                        string hprice = itemObj["hprice"].ToString();//최고가 정보이다. 최고가 정보가 없거나 가격비교 데이터가 없는 경우 0으로 표시된다.
                        string mallName = itemObj["mallName"].ToString();//상품을 판매하는 쇼핑몰의 상호이다. 정보가 없을 경우 네이버로 표기된다.
                        string productId = itemObj["productId"].ToString();//해당 상품에 대한 ID 이다.
                        string productType = itemObj["productType"].ToString();//상품군 정보를 일반상품, 중고상품, 단종상품, 판매예정상품으로 구분한다.


                        if (lprice.Equals(""))
                        {
                            lprice = "0";
                        }

                        if (hprice.Equals(""))
                        {
                            hprice = "0";
                        }

                        title = title.Replace("<b>", "").Replace("</b>", "");
                        List<string> titleKeywordList = title.Split(' ').ToList();
                        titleKeywordListResult.AddRange(titleKeywordList);
                        ShopResult.Add(new ShopAPIResult
                        {
                            RelKeyword = keyWord,
                            Title = title,
                            Image = image,
                            Lprice = lprice,
                            Hprice = hprice,
                            MallName = mallName,
                            ProductId = productId,
                            ProductType = productType,
                            Link = link,
                            TitleKeywordList = titleKeywordList
                        });
                        roop++;
                    }//검색 데이터 다수 END
                } // json[items] roop END 

            } // 물품 다수 END

            Result = (new OpenApiDataSetResult { TitleKeywordList = titleKeywordListResult, ShopAPIResultList = ShopResult, Total = total });
            return Result;
        }


        /// <summary>
        /// 쇼핑연관 키워드와 상품수를 가져온다.
        /// </summary>
        /// <param name="KeyWord"></param>
        /// <param name="TotalProdutCount"></param>
        /// <returns></returns>
        public List<KeywordList> ShopRelKeyword(string KeyWord, out int TotalProdutCount)
        {
            int totalNo = 0;
            string tempProductSet_total = "0";

            string url = "https://search.shopping.naver.com/search/all.nhn?origQuery=" + KeyWord + "&pagingIndex=" + Convert.ToString(1) + "&pagingSize=40&viewType=list&sort=rel&frm=NVSHPAG&query=" + KeyWord;
            List<KeywordList> Datalist = new List<KeywordList>();
            string textHtml = HttpWebRequestText(url);
            string outtext = string.Empty;
            agi.HtmlDocument doc = new agi.HtmlDocument();
            doc.LoadHtml(textHtml);
            var htmlNode = doc.DocumentNode.SelectSingleNode("//*[@id=\"__NEXT_DATA__\"]");
            if (htmlNode != null)
            {
                string jsonDataset = htmlNode.InnerHtml;
                JObject obj = JObject.Parse(jsonDataset);
                JObject props = JObject.Parse(obj["props"].ToString());
                JObject pageProps = JObject.Parse(props["pageProps"].ToString());

                if (pageProps["tags"] != null)
                {
                    JArray array = JArray.Parse(pageProps["tags"].ToString());
                    foreach (JObject item in array)
                    {
                        string refKeyword = item["tagName"].ToString();
                        ReturnToLabel(refKeyword);
                        Datalist.Add(new KeywordList() { Keyword = refKeyword, Kind = "S" });
                    }
                }


                JObject initialState = JObject.Parse(pageProps["initialState"].ToString());
                JObject products = JObject.Parse(initialState["products"].ToString());


                if (products["total"] != null)
                {
                    tempProductSet_total = products["total"].ToString();
                }

                totalNo = Convert.ToInt32(tempProductSet_total);
                TotalProdutCount = totalNo;
            }
            else
            {
                TotalProdutCount = 0;
            }

            return Datalist;
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
            return responseText;
        }
    }
}
