using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.IO;
using System.Net;
using System.Threading;
using System.Windows.Forms;
using agi = HtmlAgilityPack;
using Newtonsoft.Json.Linq;

namespace keywordGOGO
{

    class bizranking
    {
        // 메인폼 전달 이벤트 선언
        public static event listBoxText ReturnToMessage;
        public static event labelText ReturnToLabel;
        private SQLiteConnection conn = null;
        /// <summary>
        /// 타겟 URL 부터 HTML 코드를 가져온다.
        /// </summary>
        /// <param name="tagetUrl">타겟 URL</param>
        /// <returns>String으로 된 HTML 소스 </returns>
        private string httpWebRequestText(string tagetUrl)
        {
            string responseText = string.Empty;
            string url = tagetUrl;
            Thread.Sleep(2000);
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);
            request.Method = "GET";
            request.Timeout = 30 * 1000; // 30초
            request.Headers.Add("Authorization", "BASIC SGVsbG8="); // 헤더 추가 방법

            using (HttpWebResponse resp = (HttpWebResponse)request.GetResponse())
            {
                HttpStatusCode status = resp.StatusCode;

                Console.WriteLine(status);  // 정상이면 "OK"

                //listBox1.Items.Add("네이버와 통신결과 : " + status);

                Stream respStream = resp.GetResponseStream();
                using (StreamReader sr = new StreamReader(respStream))
                {
                    responseText = sr.ReadToEnd();
                }
            }
            return responseText;
        }


        public void JSONParser(string textHtml)
        {
            agi.HtmlDocument doc = new agi.HtmlDocument();
            doc.LoadHtml(textHtml);
            var htmlNode = doc.DocumentNode.SelectSingleNode("//*[@id=\"__NEXT_DATA__\"]");
            string jsonDataset = htmlNode.InnerHtml;
            JObject obj = JObject.Parse(jsonDataset);
        }



        /// <summary>
        /// 조회된 페이지의 HTML을 분석한다.
        /// </summary>
        /// <param name="textHtml">원본 html 데이터</param>
        /// <returns>상품리스트의 정보를 딕셔너리에 담아 리스트화 한다.</returns>
        public List<Dictionary<string, string>> HTMLParser(string textHtml, int pageNo, string keyword)
        {
            string outtext = string.Empty;
            agi.HtmlDocument doc = new agi.HtmlDocument();
            doc.LoadHtml(textHtml);
            IList<agi.HtmlNode> nodes = doc.DocumentNode.SelectNodes("//*[@id=\"__next\"]/div/div[2]/div/div[3]/div[1]/ul/li");
            List<Dictionary<string, string>> dataList = new List<Dictionary<string, string>>();

            int count = 1;
            int rank = 1;
            foreach (var node in nodes)
            {

                Dictionary<string, string> dicData = new Dictionary<string, string>();

                var naverArea = node.Attributes["class"].Value; // 조회 데이터 영역정보
                var productNo = "";//node.Attributes["data-nv-mid"].Value; // 상품번호
                var classInfo = node.Attributes["class"].Value; // 조회데이터 클래스 정보 
                var productName = node.SelectSingleNode("div/div[2]/div[1]/a");//상품명 
                var productUrl = node.SelectSingleNode("div/div[2]/div[1]/a").Attributes["href"].Value; //상품url
                var productPrice = node.SelectSingleNode("div/div[2]/div[2]/strong/span"); //상품가격
                var mallName = node.SelectSingleNode("div/div[3]/div[1]/a[1]"); //쇼핑몰명

                string categoryName = string.Empty;

                IList<agi.HtmlNode> categoryNodes = node.SelectNodes("div/div[2]/div[3]/a");
                foreach (var categoryNode in categoryNodes)
                {
                    categoryName = categoryNode.Attributes["href"].Value; // 카테고리 데이터
                }


                dicData.Add("count", Convert.ToString(count));
                dicData.Add("Keyword", keyword);
                dicData.Add("productNo", productNo);
                dicData.Add("pageNo", Convert.ToString(pageNo));
                dicData.Add("naverArea", naverArea);
                dicData.Add("productUrl", productUrl);
                dicData.Add("classInfo", classInfo);
                dicData.Add("productName", productName.InnerText.Replace("\n", "").Trim());
                dicData.Add("productPrice", productPrice.InnerText.Replace("\n", "").Trim());
                dicData.Add("rank", Convert.ToString(rank));


                ReturnToLabel(productName.InnerText.Replace("\n", "").Trim());

                if (mallName != null)
                {
                    dicData.Add("mallName", mallName.InnerText.Replace("\n", "").Trim());
                }
                else
                {
                    dicData.Add("mallName", "가격비교");
                }

                dicData.Add("categoryName", categoryName.Replace("/search/category?catId=", ""));

                /*
                Console.WriteLine(Convert.ToString(count));
                Console.WriteLine(productUrl);
                Console.WriteLine(productName.InnerText.Replace("\n","").Trim());
                Console.WriteLine(naverArea);
                Console.WriteLine(classInfo);
                Console.WriteLine(productPrice.InnerText.Replace("\n", "").Trim());
                Console.WriteLine(mallName.InnerText.Replace("\n", "").Trim());
                Console.WriteLine(categoryName.Replace("cat_id_", ""));
                */
                rank++;
                count++;
                dataList.Add(dicData);
            }

            return dataList;
        }

        public int totalProdutCount(string textHtml)
        {
            int totalNo = 0;
            string tempProductSet_total = "0";
            string outtext = string.Empty;
            agi.HtmlDocument doc = new agi.HtmlDocument();
            doc.LoadHtml(textHtml);
            var _productSet_total = doc.DocumentNode.SelectSingleNode("//*[@id=\"__next\"]/div/div[2]/div/div[3]/div[1]/div[1]/ul/li[1]/button/span[1]");
            if (_productSet_total != null)
            {
                tempProductSet_total = Convert.ToString(_productSet_total.InnerText).Replace(",", "").Replace("전체", "").Trim(); ;
            }

            totalNo = Convert.ToInt32(tempProductSet_total);
            return totalNo;
        }

        public GridResultData2 SamartStoreRankingSearch(string keyword, string mallNamedata, string adpricedata)
        {
            string strConn = @"Data Source=" + Application.StartupPath + "\\apiQc.db";
            conn = new SQLiteConnection(strConn);
            conn.Open();

            GridResultData2 Result = new GridResultData2();
            List<Dictionary<string, string>> resultDataList = new List<Dictionary<string, string>>();
            List<Dictionary<string, string>> tempDataList = new List<Dictionary<string, string>>();

            List<RankingList> ADResult = new List<RankingList>();
            List<RankingList> NonADResult = new List<RankingList>();

            string countUrl = "https://search.shopping.naver.com/search/all.nhn?query=" + keyword + "&frm=NVSCVUI";
            string countHtml = httpWebRequestText(countUrl);
            int product_totoal = totalProdutCount(countHtml);

            ReturnToMessage("상품을 " + Convert.ToString(product_totoal) + "개 찾았습니다.");
            //Console.WriteLine(Convert.ToString(product_totoal));

            int lastPageNo = 0;

            if (product_totoal > 0)
            {
                double dTotal = product_totoal / 40;
                if (Math.Ceiling(dTotal) >= 10)
                {
                    lastPageNo = 10;
                }
                else
                {
                    lastPageNo = Convert.ToInt32(Math.Ceiling(dTotal));
                }



                for (int pageNo = 1; pageNo <= lastPageNo; pageNo++)
                {
                    string url = "https://search.shopping.naver.com/search/all.nhn?origQuery=" + keyword + "&pagingIndex=" + Convert.ToString(pageNo) + "&pagingSize=40&viewType=list&sort=rel&frm=NVSHPAG&query=" + keyword;

                    //Thread.Sleep(1000);
                    string textHtml = httpWebRequestText(url);

                    tempDataList = HTMLParser(textHtml, pageNo, keyword);
                    JSONParser(textHtml);
                    resultDataList.AddRange(tempDataList);

                }


                foreach (Dictionary<string, string> resultDic in resultDataList)
                {
                    string count = resultDic["count"];
                    string naverArea = resultDic["naverArea"];
                    string productNo = resultDic["productNo"];
                    string classInfo = resultDic["classInfo"];
                    string productUrl = resultDic["productUrl"];
                    string productPrice = resultDic["productPrice"];
                    string productName = resultDic["productName"];
                    string mallName = resultDic["mallName"];
                    string categoryName = resultDic["categoryName"];
                    string pageNo = resultDic["pageNo"];
                    string reKeyword = resultDic["Keyword"];
                    string rank = resultDic["rank"];

                    ReturnToLabel(productName);

                    if (classInfo.IndexOf("ad") > 0)
                    {
                        if (mallName == mallNamedata)
                        {

                            string sqlFormattedDate = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                            String sql1 = "insert into adrank (mallname, keyword,rank,page,price,productNo,date) values('" + mallName + "','" + reKeyword + "','" + rank + "','" + pageNo + "','" + adpricedata + "','" + productNo + "','" + sqlFormattedDate + "')";
                            SQLiteCommand command = new SQLiteCommand(sql1, conn);
                            int result = command.ExecuteNonQuery();
                            Console.WriteLine("--------------------------------------");
                            Console.WriteLine("데이터베이스 입력결과: " + Convert.ToString(result));

                            string oldRank = string.Empty;
                            string oldprice = string.Empty;
                            string oldPage = string.Empty;
                            string sql2 = "select* from adrank where adrank.productNo ='" + productNo + "' and adrank.mallname ='" + mallName + "' order by date desc limit 2";
                            SQLiteCommand cmd = new SQLiteCommand(sql2, conn);
                            SQLiteDataReader rdr = cmd.ExecuteReader();

                            int idxRank = 0;

                            while (rdr.Read())
                            {
                                if (idxRank == 1)
                                {
                                    oldRank = Convert.ToString(rdr["rank"]);
                                    oldprice = Convert.ToString(rdr["price"]);
                                    oldPage = Convert.ToString(rdr["page"]);
                                }

                                idxRank++;
                            }

                            ADResult.Add(new RankingList() { rank = Convert.ToString(rank), pageNo = pageNo, productNo = productNo, oldPageNo = oldPage, count = count, mallName = mallName, productName = productName, keyword = reKeyword, productUrl = productUrl, productPrice = productPrice, categoryName = categoryName, oldrank = oldRank, adprice = adpricedata, oldadprice = oldprice });
                        }
                    }
                    else
                    {
                        if (mallName == mallNamedata)
                        {
                            string sqlFormattedDate = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                            String sql1 = "insert into nadrank (mallname, keyword,rank,page,productNo,date) values('" + mallName + "','" + reKeyword + "','" + rank + "','" + pageNo + "','" + productNo + "','" + sqlFormattedDate + "')";
                            SQLiteCommand command = new SQLiteCommand(sql1, conn);
                            int result = command.ExecuteNonQuery();
                            Console.WriteLine("--------------------------------------");
                            Console.WriteLine("데이터베이스 입력결과: " + Convert.ToString(result));

                            string oldRank = string.Empty;
                            string oldprice = string.Empty;
                            string oldPage = string.Empty;
                            string sql2 = "select* from nadrank where nadrank.productNo ='" + productNo + "' and nadrank.mallname ='" + mallName + "' order by date desc limit 2";
                            SQLiteCommand cmd = new SQLiteCommand(sql2, conn);
                            SQLiteDataReader rdr = cmd.ExecuteReader();

                            int idxRank = 0;

                            while (rdr.Read())
                            {
                                if (idxRank == 1)
                                {
                                    oldRank = Convert.ToString(rdr["rank"]);
                                    oldPage = Convert.ToString(rdr["page"]);

                                }

                                idxRank++;
                            }

                            NonADResult.Add(new RankingList() { rank = Convert.ToString(rank), pageNo = pageNo, oldPageNo = oldPage, productNo = productNo, count = count, mallName = mallName, productName = productName, keyword = reKeyword, productUrl = productUrl, productPrice = productPrice, categoryName = categoryName, oldrank = oldRank, adprice = "-" });
                        }
                    }
                }
            }
            ReturnToMessage("데이터를 출력합니다.");
            Result = new GridResultData2() { AdRankingRefGrid = ADResult, NonAdRankingRefGrid = NonADResult };
            conn.Close();
            return Result;

        }

    }
}
