using System;
using agi = HtmlAgilityPack;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using System.Net;
using System.IO;
using System.Windows.Forms;
using System.Threading;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data.SQLite;


namespace keywordGOGO
{
    
    class bizranking
    {
        // 메인폼 전달 이벤트 선언
        public static event listBoxText ReturnToMessage;
        public static event labelText ReturnToLabel;

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
            IList<agi.HtmlNode> nodes = doc.QuerySelectorAll("#_search_list > div.search_list.basis > ul >li");
            List<Dictionary<string, string>> dataList = new List<Dictionary<string, string>>();

            int count = 1;
            foreach (var node in nodes)
            {
                Dictionary<string, string> dicData = new Dictionary<string, string>();

                var naverArea = node.Attributes["data-expose-area"].Value; // 조회 데이터 영역정보
                var classInfo = node.Attributes["class"].Value; // 조회데이터 클래스 정보 
                var productName = node.QuerySelector("div.info > div > a"); //상품명
                var productUrl = node.QuerySelector("div.info > div > a").Attributes["href"].Value; //상품url
                var productPrice = node.QuerySelector("div.info > span.price > em > span"); //상품가격
                var mallName = node.QuerySelector("div.info_mall > p > a.mall_img"); //쇼핑몰명

                string categoryName = string.Empty;

                IList<agi.HtmlNode> categoryNodes = node.QuerySelectorAll("div.info > span.depth > a");
                foreach (var categoryNode in categoryNodes)
                {
                    categoryName = categoryNode.Attributes["class"].Value; // 카테고리 데이터
                }


                dicData.Add("count", Convert.ToString(count));
                dicData.Add("Keyword", keyword);
                dicData.Add("pageNo", Convert.ToString(pageNo));
                dicData.Add("naverArea", naverArea);
                dicData.Add("productUrl", productUrl);
                dicData.Add("classInfo", classInfo);
                dicData.Add("productName", productName.InnerText.Replace("\n", "").Trim());
                dicData.Add("productPrice", productPrice.InnerText.Replace("\n", "").Trim());

                if (mallName != null)
                {
                    dicData.Add("mallName", mallName.InnerText.Replace("\n", "").Trim());
                }
                else
                {
                    dicData.Add("mallName", "가격비교");
                }

                dicData.Add("categoryName", categoryName.Replace("cat_id_", ""));

                /**
                Console.WriteLine(Convert.ToString(count));
                Console.WriteLine(productUrl);
                Console.WriteLine(productName.InnerText.Replace("\n","").Trim());
                Console.WriteLine(naverArea);
                Console.WriteLine(classInfo);
                Console.WriteLine(productPrice.InnerText.Replace("\n", "").Trim());
                Console.WriteLine(mallName.InnerText.Replace("\n", "").Trim());
                Console.WriteLine(categoryName.Replace("cat_id_", ""));
                **/

                count++;
                dataList.Add(dicData);
            }

            return dataList;
        }



    }
}
