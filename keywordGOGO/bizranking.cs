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


        public List<Dictionary<string, string>> JSONParser(string textHtml, int pageNo, string keyword)
        {
            List<Dictionary<string, string>> dataList = new List<Dictionary<string, string>>();
            int count = 1;
            int rank = 1;

            agi.HtmlDocument doc = new agi.HtmlDocument();
            doc.LoadHtml(textHtml);
            var htmlNode = doc.DocumentNode.SelectSingleNode("//*[@id=\"__NEXT_DATA__\"]");
            string jsonDataset = htmlNode.InnerHtml;
            JObject obj = JObject.Parse(jsonDataset);
            JObject props = JObject.Parse(obj["props"].ToString());
            JObject pageProps = JObject.Parse(props["pageProps"].ToString());
            JObject initialState = JObject.Parse(pageProps["initialState"].ToString());
            JObject products = JObject.Parse(initialState["products"].ToString());
            JArray array = JArray.Parse(products["list"].ToString());
            foreach (JObject item in array)
            {
                Dictionary<string, string> dicData = new Dictionary<string, string>();

                JObject productitem = JObject.Parse(item["item"].ToString());
                string naverArea = "";// 클래스 정보
                string classInfo = "";
                string productNo = "";
                string mallName = "";
                
                if (productitem["mallProductId"] != null)
                {
                    productNo = productitem["mallProductId"].ToString(); // 상품번호
                    mallName = productitem["mallName"].ToString();// 몰네임
                }
                else
                {
                    productNo = "";
                    mallName = "";
                }

                if(productitem["adId"] != null)
                {
                    classInfo = productitem["adId"].ToString();
                }
                else
                {
                    classInfo = "";
                }

                if (productitem["mallProductUrl"] != null)
                {
                    string url = productitem["mallProductUrl"].ToString(); //주소
                    naverArea = url;
                }

                string productUrl = naverArea; //상품주소


                string productPrice = productitem["price"].ToString(); // 상품가격
                string productName = productitem["productName"].ToString();// 상품명

                string categoryName = "";
                if (productitem["category3Name"] != null)
                {
                    categoryName = productitem["category3Name"].ToString();// 카테고리

                }

                string relevance = "";
                if (productitem["relevance"] != null)
                {
                    relevance = productitem["relevance"].ToString();
                    double relevanceRatio = Convert.ToDouble(relevance) * 100;
                    relevance = Convert.ToString(relevanceRatio);
                    Console.WriteLine(relevanceRatio);
                }
                else
                {
                    relevance = "";
                }

                string similarity = "";
                if (productitem["similarity"] != null)
                {
                    similarity = productitem["similarity"].ToString();
                    double similarityRatio = Convert.ToDouble(similarity)*100;
                    similarity = Convert.ToString(similarityRatio);
                    Console.WriteLine(similarityRatio);
                }
                else
                {
                    similarity = "";
                }

                string hitRank = "";
                if (productitem["rank"] != null)
                {
                    hitRank = productitem["rank"].ToString();
                }
                else
                {
                    hitRank = "";
                }

                Console.WriteLine("++++++++++++++++++++++++++");
                Console.WriteLine(rank);
                Console.WriteLine(keyword);
                Console.WriteLine(productPrice);
                Console.WriteLine(productPrice);

                dicData.Add("count", Convert.ToString(count));
                dicData.Add("Keyword", keyword);
                dicData.Add("productNo", productNo); // 상품번호
                dicData.Add("pageNo", Convert.ToString(pageNo)); //페이지번호
                dicData.Add("naverArea", naverArea);
                dicData.Add("productUrl", productUrl);
                dicData.Add("classInfo", classInfo);
                dicData.Add("productName", productName.Replace("\n", "").Trim());
                dicData.Add("productPrice", productPrice.Replace("\n", "").Trim());
                dicData.Add("rank", Convert.ToString(rank));
                dicData.Add("relevance", Convert.ToString(relevance));
                dicData.Add("similarity", Convert.ToString(similarity));
                dicData.Add("hitRank", Convert.ToString(hitRank));


                ReturnToLabel(productName.Replace("\n", "").Trim());

                if (mallName != null)
                {
                    dicData.Add("mallName", mallName.Replace("\n", "").Trim());
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

            var htmlNode = doc.DocumentNode.SelectSingleNode("//*[@id=\"__NEXT_DATA__\"]");
            string jsonDataset = htmlNode.InnerHtml;
            JObject obj = JObject.Parse(jsonDataset);
            JObject props = JObject.Parse(obj["props"].ToString());
            JObject pageProps = JObject.Parse(props["pageProps"].ToString());
            JObject initialState = JObject.Parse(pageProps["initialState"].ToString());
            JObject products = JObject.Parse(initialState["products"].ToString());


            if (products["total"] != null)
            {
                tempProductSet_total = products["total"].ToString();
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
            List<RankingList> AllResult = new List<RankingList>();

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

                    //tempDataList = HTMLParser(textHtml, pageNo, keyword);
                    tempDataList = JSONParser(textHtml,pageNo, keyword);
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
                    if(string.IsNullOrEmpty(productNo))
                    {
                        productNo = "가격비교상품";
                    }
                    string categoryName = resultDic["categoryName"];
                    string pageNo = resultDic["pageNo"];
                    string reKeyword = resultDic["Keyword"];
                    string rank = resultDic["rank"];

                    string relevance = resultDic["relevance"];
                    string similarity = resultDic["similarity"];
                    string hitRank = resultDic["hitRank"];
                    string adyn = "";

                    ReturnToLabel(productName);

                   

                    if (classInfo.IndexOf("ad") > 0)
                    {
                        adyn = "광고상품";
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

                            ADResult.Add(new RankingList() { rank = Convert.ToString(rank), pageNo = pageNo, productNo = productNo, oldPageNo = oldPage, count = count, mallName = mallName, productName = productName, keyword = reKeyword, productUrl = productUrl, productPrice = productPrice, categoryName = categoryName, oldrank = oldRank, adprice = adpricedata, oldadprice = oldprice, similarity = similarity, relevance= relevance, hitRank= hitRank });
                        }
                    }
                    else
                    {
                        adyn = "일반상품";
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

                            NonADResult.Add(new RankingList() { rank = Convert.ToString(rank), pageNo = pageNo, oldPageNo = oldPage, productNo = productNo, count = count, mallName = mallName, productName = productName, keyword = reKeyword, productUrl = productUrl, productPrice = productPrice, categoryName = categoryName, oldrank = oldRank, adprice = "-", similarity = similarity, relevance = relevance, hitRank = hitRank});
                        }
                    }

                    AllResult.Add(new RankingList() { rank = Convert.ToString(rank), pageNo = pageNo, productNo = productNo, count = count, mallName = mallName, productName = productName, keyword = reKeyword, productUrl = productUrl, productPrice = productPrice, categoryName = categoryName, adprice = "-", similarity = similarity, relevance = relevance, hitRank = hitRank, adYn = adyn });
                }
            }
            ReturnToMessage("데이터를 출력합니다.");
            Result = new GridResultData2() { AdRankingRefGrid = ADResult, NonAdRankingRefGrid = NonADResult, AllRankingRefGrid = AllResult };
            conn.Close();
            return Result;

        }

    }
}
