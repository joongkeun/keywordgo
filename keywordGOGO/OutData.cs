using System.Net;
using System.Threading;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using System.Text.RegularExpressions;
using System.Data.SQLite;
using RestSharp;
using System.IO;
using System.Windows.Forms;

namespace keywordGOGO
{
    class OutData
    {
        // 메인폼 전달 이벤트 선언
        public static event listBoxText ReturnToMessage;
        public static event labelText ReturnToLabel;
        public GridResultData GridDataSet(string KeyWord, int RefMaxCount, SQLiteConnection conn)
        {
            int sellPrdCnt = 0;
            int apiUseCount = 0;
            int refAdTotalQcCnt = 0;
            GridResultData Result = new GridResultData();
            List<KeyWordResult> ADResult = new List<KeyWordResult>();
            List<KeyWordResult> RefShopResult = new List<KeyWordResult>();
            List<KeyWordResult> SEOResult = new List<KeyWordResult>();
            List<KeywordList> RefShopKeyWord = new List<KeywordList>();
            List<ShopAPIResult> OpenApiList = new List<ShopAPIResult>();
            List<String> TitleKeywordList = new List<string>();
            List<ProductKeyWordList> ProductWordList = new List<ProductKeyWordList>();


            OpenApiDataSetResult openApiDataSetResult = new OpenApiDataSetResult();
            NaverApi naverApi = new NaverApi();
            NaverShoppingCrawler shoppingCrawler = new NaverShoppingCrawler();

            ReturnToMessage("광고 연관 검색어를 조회합니다..");

            // 네이버 광고 API 실행
            string naver = naverApi.NaverAdApi(KeyWord.Replace(" ", ""));
            JObject obj = JObject.Parse(naver);
            JArray array = JArray.Parse(obj["keywordList"].ToString());

            // 네이버 키워드 도구 연관 검색어(조회키워드만)
            foreach (JObject itemObj in array)
            {
                string relKeyword = itemObj["relKeyword"].ToString();//키워드
                string monthlyPcQcCnt = itemObj["monthlyPcQcCnt"].ToString();//월간 pc 검색수 
                string monthlyMobileQcCnt = itemObj["monthlyMobileQcCnt"].ToString();//월간 모바일 검색수
                string monthlyAvePcClkCnt = itemObj["monthlyAvePcClkCnt"].ToString();//월간 PC 클릭수
                string monthlyAveMobileClkCnt = itemObj["monthlyAveMobileClkCnt"].ToString();//월간 모바일 클릭수
                string monthlyAvePcCtr = itemObj["monthlyAvePcCtr"].ToString();//월간 PC 클릭률
                string monthlyAveMobileCtr = itemObj["monthlyAveMobileCtr"].ToString();//월간 모바일 클릭률
                string plAvgDepth = itemObj["plAvgDepth"].ToString();//경쟁정도
                string compIdx = itemObj["compIdx"].ToString();// 월간노출광고수

                monthlyPcQcCnt = monthlyPcQcCnt.Replace("<", "");
                monthlyMobileQcCnt = monthlyMobileQcCnt.Replace("<", "");



                ReturnToLabel(relKeyword);

                // 네이버 쇼핑 연관 검색어 + 네이버 키워드 도구 연관 검색어
                openApiDataSetResult = naverApi.OpenApiDataSet(relKeyword); // 오픈API를 상품을 조회한다.
                sellPrdCnt = openApiDataSetResult.Total; // 키워드의 전체 상품수
                ADResult.Add(new KeyWordResult() {RelKeyword = relKeyword, MonthlyPcQcCnt = monthlyPcQcCnt, MonthlyMobileQcCnt = monthlyMobileQcCnt, MonthlyAvePcClkCnt = monthlyAvePcClkCnt, MonthlyAveMobileClkCnt = monthlyAveMobileClkCnt, MonthlyAvePcCtr = monthlyAvePcCtr , MonthlyAveMobileCtr = monthlyAveMobileCtr, PlAvgDepth = plAvgDepth, CompIdx = compIdx, SellPrdQcCnt = sellPrdCnt, ShopResult = openApiDataSetResult.ShopAPIResultList });

                apiUseCount++; // 오픈 api 사용량 체크

                // 검색량 조절
                if(RefMaxCount <= apiUseCount)
                {
                    // 조회를 멈춘다.
                    break;
                }

            }

            // api 사용량 DB에 저장
            string sqlFormattedDate = DateTime.Now.ToString("yyyy-MM-dd");
            string sql2 = "select* from apicount where apicount.date ='" + sqlFormattedDate + "'";
            SQLiteCommand cmd = new SQLiteCommand(sql2, conn);
            SQLiteDataReader rdr = cmd.ExecuteReader();

            int idxRank = 0;
            string count = string.Empty;
            while (rdr.Read())
            {
                if (idxRank == 0)
                {
                    count = Convert.ToString(rdr["count"]);
                }

                idxRank++;
            }

            Console.WriteLine("데이터베이스 출력 결과: " + count);

            if (string.IsNullOrEmpty(count))
            {
                String sql = "insert into apicount (date,count) values('" + sqlFormattedDate + "','" + apiUseCount + "')";
                SQLiteCommand command = new SQLiteCommand(sql, conn);
                int result = command.ExecuteNonQuery();
                Console.WriteLine("--------------------------------------");
                Console.WriteLine("데이터베이스 입력결과: " + Convert.ToString(result));
            }
            else
            {
                int apiTotal = Convert.ToInt32(count) + apiUseCount;
                String sql = "update apicount set count='" + apiTotal + "'where date ='" + sqlFormattedDate + "'";
                SQLiteCommand command = new SQLiteCommand(sql, conn);

                int result = command.ExecuteNonQuery();
                Console.WriteLine("--------------------------------------");
                Console.WriteLine("데이터베이스 입력결과: " + Convert.ToString(result));
            }



            Result = new GridResultData() { RefAdTotalQcCnt = refAdTotalQcCnt, AdRefGrid = ADResult };
            return Result;
        }

        public GridResultData SubGridDataSet(string KeyWord, bool tagYn)
        {
            GridResultData Result = new GridResultData();

            //List<KeyWordResult> AllResult = new List<KeyWordResult>();
            //List<KeyWordResult> RefResult = new List<KeyWordResult>();
            List<KeywordList> RefShopKeyWord = new List<KeywordList>();
            ///List<ShopAPIResult> OpenApiList = new List<ShopAPIResult>();
            ///List<String> TitleKeywordList = new List<string>();

            //OpenApiDataSetResult openApiDataSetResult = new OpenApiDataSetResult();

            NaverApi naverApi = new NaverApi();
            NaverShoppingCrawler shoppingCrawler = new NaverShoppingCrawler();

            // 쇼핑 연관검색어와 SEO 태그를 조회한다.
            ReturnToMessage("쇼핑연관 검색어를 조회중입니다.");

            // 네이버 쇼핑 연관 검색어 + 태그 정보
            RefShopKeyWord = naverApi.ShopRelKeyword(KeyWord, out int TotalProdutCount);
            ShopWebResult webResult = new ShopWebResult();
            if (tagYn == true) { 
                ReturnToMessage("태그정보를 분석합니다.");
                webResult = shoppingCrawler.SmartStoreInfoFinder(KeyWord);
            }
            else
            {
                
                ReturnToMessage("SEO 태그 분석을 제외하였습니다.");
            }

            Console.WriteLine("전 : " + RefShopKeyWord.Count);

            Result = new GridResultData() { ShoppingRefGrid = RefShopKeyWord, ShopWebDataResult = webResult };

            return Result;

        }

    }
}
