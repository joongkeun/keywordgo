using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Threading;
using agi = HtmlAgilityPack;
using Newtonsoft.Json.Linq;
using Microsoft.Office.Interop.Excel;


namespace keywordGOGO
{
    class SaleAmount
    {
        // 메인폼 전달 이벤트 선언
        //public static event listBoxText ReturnToMessage;
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

                            string leadTimeCount1 = "0";
                            string leadTimeCount2 = "0";
                            string leadTimeCount3 = "0";
                            string leadTimeCount4 = "0";
                            string totalReviewCount = "0";
                            string averageReviewScore = "0";
                            int total = 0;
                            string jsonDataset = htmlNode.InnerHtml;
                            try
                            {
                                JObject obj = JObject.Parse(jsonDataset.Replace("window.__PRELOADED_STATE__=", ""));
                                JObject product = JObject.Parse(obj["product"].ToString());
                                JObject A = JObject.Parse(product["A"].ToString());


                                if (A["productDailyDeliveryLeadTimes"] != null)
                                {
                                    JObject productDeliveryLeadTimes = JObject.Parse(A["productDailyDeliveryLeadTimes"].ToString());

                                    leadTimeCount1 = productDeliveryLeadTimes["leadTimeCount"][0].ToString();
                                    leadTimeCount2 = productDeliveryLeadTimes["leadTimeCount"][1].ToString();
                                    leadTimeCount3 = productDeliveryLeadTimes["leadTimeCount"][2].ToString();
                                    leadTimeCount4 = productDeliveryLeadTimes["leadTimeCount"][3].ToString();
                                }
                                else
                                {
                                    leadTimeCount1 = "0";
                                    leadTimeCount2 = "0";
                                    leadTimeCount3 = "0";
                                    leadTimeCount4 = "0";
                                }

                                if (A["reviewAmount"] != null)
                                {

                                    JObject reviewAmount = JObject.Parse(A["reviewAmount"].ToString());
                                    totalReviewCount = reviewAmount["totalReviewCount"].ToString();
                                    averageReviewScore = reviewAmount["averageReviewScore"].ToString();

                                }
                                else
                                {
                                    totalReviewCount = "0";
                                    averageReviewScore = "0";
                                }
                            

                            //string recentSaleCount = productDeliveryLeadTimes["recentSaleCount"].ToString();

                                 total = int.Parse(leadTimeCount1) + int.Parse(leadTimeCount2) + int.Parse(leadTimeCount3) + int.Parse(leadTimeCount4);
                            }
                            catch (Exception ex)
                            {
                                leadTimeCount1 = "parsing error";
                                leadTimeCount2 = "parsing error";
                                leadTimeCount3 = "parsing error";
                                leadTimeCount4 = "parsing error";
                                totalReviewCount = "parsing error";
                                averageReviewScore = "parsing error";
                            }
                            saleAmountResults.Add(new SaleAmountResult()
                                {
                                    productName = productName, // 상품명
                                    mallName = mallName, // 몰이름
                                    categoryName = categoryText, // 카테고리명
                                    openDate = openDate, // 오픈일
                                    totalReviewCount = totalReviewCount, // 리뷰수
                                    averageReviewScore = averageReviewScore,// 평점
                                    leadTimeCount1 = leadTimeCount1,
                                    leadTimeCount2 = leadTimeCount2,
                                    leadTimeCount3 = leadTimeCount3,
                                    leadTimeCount4 = leadTimeCount4, 
                                    totalleadTimeCount1 = total.ToString() ,
                                    //cumulationSaleCount = cumulationSaleCount,//6개월 판매수
                                    //recentSaleCount = recentSaleCount, //최근 3일 판매수
                                    urlLink = productUrl

                                });

                                ReturnToLabel(mallName);

                                rowidx++;

                            
                        }
                    }
                }
            }

            //중복 제거


            return saleAmountResults;
        }

    }
}
