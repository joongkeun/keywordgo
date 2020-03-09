using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using System.Text.RegularExpressions;
using System.Data.SQLite;
using RestSharp;
using System.IO;
using System.Windows.Forms;
using agi = HtmlAgilityPack;

namespace keywordGOGO
{
    class NaverShoppingCrawler
    {
        // 메인폼 전달 이벤트 선언
        public static event listBoxText ReturnToMessage;
        public static event labelText ReturnToLabel;

        /// <summary>
        /// 조회된 페이지의 HTML을 분석한다.
        /// </summary>
        /// <param name="textHtml">원본 html 데이터</param>
        /// <returns>상품리스트의 정보를 딕셔너리에 담아 리스트화 한다.</returns>
        public List<Dictionary<string, string>> HTMLParser(string textHtml, int pageNo, string keyword, out int adCnt)
        {
            int addT = 0;
            string outtext = string.Empty;
            agi.HtmlDocument doc = new agi.HtmlDocument();
            doc.LoadHtml(textHtml);
            IList<agi.HtmlNode> nodes = doc.QuerySelectorAll("#_search_list > div.search_list.basis > ul >li");
            List<Dictionary<string, string>> dataList = new List<Dictionary<string, string>>();

            int count = 1;
            foreach (var node in nodes)
            {
                Dictionary<string, string> dicData = new Dictionary<string, string>();
                string smartFarmYn = string.Empty;
                var naverArea = node.Attributes["data-expose-area"].Value; // 조회 데이터 영역정보
                var classInfo = node.Attributes["class"].Value; // 조회데이터 클래스 정보 
                var productName = node.QuerySelector("div.info > div > a"); //상품명
                var productUrl = node.QuerySelector("div.info > div > a").Attributes["href"].Value; //상품url
                var productPrice = node.QuerySelector("div.info > span.price > em > span"); //상품가격
                var mallName = node.QuerySelector("div.info_mall > p > a.mall_img"); //쇼핑몰명


                if (classInfo == "ad _itemSection")
                {
                    addT++;
                }

                if (naverArea == "lst*N" || naverArea == "lst*A")
                {
                    smartFarmYn = node.Attributes["data-is-shop-n"].Value; // 스마트팜유무
                }

                string categoryName = string.Empty;
                string categoryText = string.Empty;
                IList<agi.HtmlNode> categoryNodes = node.QuerySelectorAll("div.info > span.depth > a");
                foreach (var categoryNode in categoryNodes)
                {
                    categoryName = categoryNode.Attributes["class"].Value; // 카테고리 데이터
                    categoryText = categoryNode.InnerText;
                }

                ReturnToLabel(keyword);

                dicData.Add("count", Convert.ToString(count)); // 1page의 상품수
                dicData.Add("Keyword", keyword); //조회키워드
                dicData.Add("pageNo", Convert.ToString(pageNo));
                dicData.Add("naverArea", naverArea);
                dicData.Add("productUrl", productUrl);
                dicData.Add("classInfo", classInfo);
                dicData.Add("productName", productName.InnerText.Replace("\n", "").Trim());
                dicData.Add("productPrice", productPrice.InnerText.Replace("\n", "").Trim());
                dicData.Add("smartFarmYn", smartFarmYn);
                dicData.Add("categoryText", categoryText);

                if (mallName != null)
                {
                    dicData.Add("mallName", mallName.InnerText.Replace("\n", "").Trim());
                }
                else
                {
                    dicData.Add("mallName", "가격비교");
                }

                dicData.Add("categoryName", categoryName.Replace("cat_id_", ""));


                Console.WriteLine(Convert.ToString(count));
                Console.WriteLine(productUrl);
                Console.WriteLine(productName.InnerText.Replace("\n", "").Trim());
                Console.WriteLine(naverArea);
                Console.WriteLine(classInfo);
                Console.WriteLine(productPrice.InnerText.Replace("\n", "").Trim());
                Console.WriteLine(mallName.InnerText.Replace("\n", "").Trim());
                Console.WriteLine(categoryName.Replace("cat_id_", ""));


                count++;
                dataList.Add(dicData);
            }

            adCnt = addT;
            return dataList;
        }

        /// <summary>
        /// 스마트 스토어 정보 크롤링
        /// </summary>
        /// <param name="relKeyword"></param>
        /// <returns></returns>
        public ShopWebResult SmartStoreInfoFinder(string relKeyword)
        {
            /**
             * ### 엑셀에 시트별로 분석 데이터 출력 ###  
             *  - 전체상품 갯수
             *  - 1page에 있는 광고를 제외한 중복제거된 스마트스토어 태그
             *  - 상품명 리스트 나열
             *  - 상품명 또는 키워드 리스트화 중복제거
             *  - 많이 쓰는 카테고리 정보
             *  - 중복제거된 쇼핑몰명
             *  - 
             **/

            List<KeywordList> tagList = new List<KeywordList>();
            List<string> mallList = new List<string>();
            List<string> productNmList = new List<string>();
            List<Dictionary<string, string>> resultDataList = new List<Dictionary<string, string>>();

            string url = "https://search.shopping.naver.com/search/all.nhn?origQuery=" + relKeyword + "&pagingIndex=" + Convert.ToString(1) + "&pagingSize=80&viewType=list&sort=rel&frm=NVSHPAG&query=" + relKeyword;

            string textHtml = HttpWebRequestText(url);
            int totalCount = totalProdutCount(textHtml);
            List<string> shoppingRefKeyWord = ShoppingKeywordHtml(textHtml);
            resultDataList = HTMLParser(textHtml, 1, relKeyword, out int adCount);

            int rowidx = 1;
            if (totalCount > 0)
            {
                foreach (Dictionary<string, string> resultDic in resultDataList)
                {
                    string count = resultDic["count"];
                    string naverArea = resultDic["naverArea"];
                    string classInfo = resultDic["classInfo"];
                    string productUrl = resultDic["productUrl"];
                    string productPrice = resultDic["productPrice"];
                    string productName = resultDic["productName"];
                    string mallName = resultDic["mallName"];
                    string categoryName = resultDic["categoryName"];
                    string pageNo = resultDic["pageNo"];
                    string Keyword = resultDic["Keyword"];
                    string smartFarmYn = resultDic["smartFarmYn"];
                    string categoryText = resultDic["categoryText"];

                    productNmList.Add(productName);
                    ReturnToLabel(productName);
                    if (smartFarmYn == "true")
                    {
                        if (classInfo != "ad _itemSection")
                        {
                            if (rowidx < 20)
                            {
                                string key = string.Empty;
                                string value = string.Empty;

                                Dictionary<string, string> dicData = new Dictionary<string, string>();

                                // 리다이렉트 상품으로 이동
                                Thread.Sleep(1000);
                                string reouttext = HttpWebRequestText(productUrl);
                                agi.HtmlDocument redoc = new agi.HtmlDocument();
                                redoc.LoadHtml(reouttext);

                                IList<agi.HtmlNode> nodes = redoc.DocumentNode.SelectNodes("//*[@class='tb_view2']");

                                if(nodes != null) {

                                    if (nodes.Count > 0)
                                    {

                                        int nodeCnt = 0;
                                        foreach (var node in nodes)
                                        {
                                            if (nodeCnt == 0)
                                            {
                                                IList<agi.HtmlNode> thNodes = nodes.QuerySelectorAll("tbody > tr > th");
                                                IList<agi.HtmlNode> tdNodes = nodes.QuerySelectorAll("tbody > tr > td");
                                                int node2Cnt = 0;
                                                foreach (var node2 in thNodes)
                                                {

                                                    key = node2.InnerText;
                                                    value = tdNodes[node2Cnt].InnerText;
                                                    dicData[key] = value;
                                                    node2Cnt++;
                                                }
                                            }
                                        }
                                    }
                                }
                                agi.HtmlNode tagNode = redoc.DocumentNode.SelectSingleNode("//*[@class='goods_tag']");


                                string tagData = string.Empty;
                                if (tagNode != null)
                                {
                                    tagData = tagNode.InnerText.Replace("\n", "").Replace("#", ",").Replace("Tag", "").Replace("				", "").Trim();
                                    List<string> tagDataList = new List<string>(tagData.Split(','));
                                    foreach (var data in tagDataList)
                                    {
                                        ReturnToLabel(data);
                                        if(data.Length > 0)
                                            tagList.Add(new KeywordList() { Keyword = data, Kind = "T" });
                                    }

                                    Console.WriteLine(tagNode.InnerText.Replace("\n", "").Replace("#", ",").Replace("Tag", "").Replace("				", "").Trim());
                                }
                                ReturnToLabel(mallName);
                                mallList.Add(mallName);
                                rowidx++;

                            }
                        }
                    }
                }
            }

            //중복 제거
            tagList = tagList.Distinct().ToList(); //SEO 태그리스트
            mallList = mallList.Distinct().ToList(); //몰 태그리스트
            // 결과전송
            ShopWebResult result = new ShopWebResult()
            {
                RelKeyword = relKeyword, // 키워드
                AdCount = adCount, // 첫페이지 광고수
                TotalCount = totalCount, // 네이버 쇼핑 상품수
                OutTagList = tagList, // SEO 태그 리스트
                ShoppingRefKeyWord = shoppingRefKeyWord, //쇼핑 연관검색어 리스트 
                MallList = mallList, // 상위 스마트 스토어명 리스트
                ProductNmList = productNmList // 상위 상품명 리스트
            };

            return result;
        }



        /// <summary>
        /// 쇼핑 검색시 전체 상품수를 조회한다.
        /// </summary>
        /// <param name="textHtml"></param>
        /// <returns></returns>
        public int totalProdutCount(string textHtml)
        {
            int totalNo = 0;
            string tempProductSet_total = "0";
            string outtext = string.Empty;
            agi.HtmlDocument doc = new agi.HtmlDocument();
            doc.LoadHtml(textHtml);

            var _productSet_total = doc.QuerySelector("#snb > ul > li.snb_all.on > a");

            if (_productSet_total != null)
            {
                tempProductSet_total = Convert.ToString(_productSet_total.InnerText).Replace(",", "").Replace("전체", "").Trim(); ;
            }

            totalNo = Convert.ToInt32(tempProductSet_total);
            return totalNo;
        }


        public List<string> ShoppingKeywordHtml(string textHtml)
        {

            List<string> Datalist = new List<string>();
            agi.HtmlDocument doc = new agi.HtmlDocument();
            doc.LoadHtml(textHtml);
            IList<agi.HtmlNode> nodes = doc.QuerySelectorAll("#_relatedTagArea > div > ul > li");
            int count = 1;
            foreach (var node in nodes)
            {
                var refKeyword = node.QuerySelector("a").InnerText; //상품명
                //Console.WriteLine(refKeyword.Trim().Replace("\n", ""));
                refKeyword = refKeyword.Trim().Replace("\n", "");
                ReturnToLabel(refKeyword);
                Datalist.Add(refKeyword);
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
