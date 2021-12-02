using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Threading;
using agi = HtmlAgilityPack;
using Newtonsoft.Json.Linq;


namespace keywordGOGO
{
    class SaleAmount
    {
        // 메인폼 전달 이벤트 선언
        public static event listBoxText ReturnToMessage;
        public static event labelText ReturnToLabel;

        /// <summary>
        /// 스마트 스토어 정보 크롤링
        /// </summary>
        /// <param name="relKeyword"></param>
        /// <returns></returns>
        public List<SaleAmountResult> SmartStoreSaleAmountFinder(string keyword,string npayType, string sort)
        {
            List<SaleAmountResult> saleAmountResults = new List<SaleAmountResult>();

            List<Dictionary<string, string>> resultDataList = new List<Dictionary<string, string>>();

            string url = 
                "https://search.shopping.naver.com/search/all.nhn?" +
                "origQuery="+ keyword +
                "&pagingIndex="+ Convert.ToString(1) + 
                "&pagingSize=80" +
                "&viewType=list" +
                "&sort=" + sort+
                "&frm=" + npayType+
                "&query=" + keyword;

            Console.WriteLine(url);

            NaverShoppingCrawler naverShoppingCrawler = new NaverShoppingCrawler();

            string textHtml = naverShoppingCrawler.HttpWebRequestText(url);
            int totalCount = naverShoppingCrawler.totalProdutCount(textHtml);

            resultDataList = naverShoppingCrawler.JSONParser(textHtml, 1, keyword, out int adCount);

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
                    string openDate = resultDic["openDate"];


                    ReturnToLabel(productName);
                    if (smartFarmYn == "true")
                    {
                        string key = string.Empty;
                        string value = string.Empty;

                        Dictionary<string, string> dicData = new Dictionary<string, string>();

                        // 리다이렉트 상품으로 이동
                        Random random = new Random();
                        int timeRange = random.Next(500, 800);
                        Thread.Sleep(timeRange);
                        Console.WriteLine(Convert.ToString(timeRange));
                        string reouttext = naverShoppingCrawler.HttpWebRequestText(productUrl);
                        agi.HtmlDocument redoc = new agi.HtmlDocument();
                        redoc.LoadHtml(reouttext);

                        var htmlNode = redoc.DocumentNode.SelectSingleNode("/html/body/script[1]/text()");
                        if (htmlNode != null)
                        {
                            string jsonDataset = htmlNode.InnerHtml;
                            JObject obj = JObject.Parse(jsonDataset.Replace("window.__PRELOADED_STATE__=", ""));
                            JObject product = JObject.Parse(obj["product"].ToString());
                            JObject A = JObject.Parse(product["A"].ToString());
                            
                            if (!string.IsNullOrEmpty(A["saleAmount"].ToString()))
                            {
                                JObject saleAmount = JObject.Parse(A["saleAmount"].ToString());

                                string cumulationSaleCount = saleAmount["cumulationSaleCount"].ToString();
                                string recentSaleCount = saleAmount["recentSaleCount"].ToString();

                                JObject reviewAmount = JObject.Parse(A["reviewAmount"].ToString());
                                string totalReviewCount = reviewAmount["totalReviewCount"].ToString();
                                string averageReviewScore = reviewAmount["averageReviewScore"].ToString();


                                saleAmountResults.Add(new SaleAmountResult()
                                {
                                    productName = productName, // 상품명
                                    mallName = mallName, // 몰이름
                                    categoryName = categoryText, // 카테고리명
                                    openDate = openDate, // 오픈일
                                    totalReviewCount = totalReviewCount, // 리뷰수
                                    averageReviewScore = averageReviewScore,// 평점
                                    cumulationSaleCount = cumulationSaleCount,//6개월 판매수
                                    recentSaleCount = recentSaleCount, //최근 3일 판매수
                                    urlLink = productUrl

                                });

                                ReturnToLabel(mallName);

                                rowidx++;

                            }
                        }
                    }
                }
            }

            //중복 제거


            return saleAmountResults;
        }

    }
}
