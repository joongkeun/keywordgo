using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Newtonsoft.Json.Linq;

namespace keywordGOGO
{
    class ExcelTOfile
    {
        // 메인폼 전달 이벤트 선언
        public static event listBoxText ReturnToMessage;
        public static event labelText ReturnToLabel;
        

        public void dataSheet(object keyWord, object SEOkeyWord, object titleKeyWord)
        {
            List<string> ExData = new List<string>();
            List<RelKeyWordResult> OutData = new List<RelKeyWordResult>();
            List<ProductKeyWordList> prdkeword =new List<ProductKeyWordList>();
            // 연관 키워드
            List<KeyWordResult> collection = keyWord as List<KeyWordResult>;
            foreach (var r in collection)
            {
                ExData.Add(r.RelKeyword);
            }


            // 상품명 키워드
            List<ProductKeyWordList> titlelist = titleKeyWord as List<ProductKeyWordList>;

            foreach (var v in titlelist)
            {
                ExData.Add(v.value);
            }

            // SEO키워드 
            List<KeywordList> seolist = SEOkeyWord as List<KeywordList>;
            foreach (var r in seolist)
            {
                ExData.Add(r.Keyword);
            }


            // 중복 단어의 수를 체크한다.
            var q = ExData.GroupBy(x => x)
           .Select(g => new { Value = g.Key, Count = g.Count() })
           .OrderByDescending(x => x.Count).ToList();
            //중복 키워드를 리스트에 담는다.
            foreach (var temp in q)
            {
                prdkeword.Add(new ProductKeyWordList() { value = temp.Value, count = temp.Count });
            }

        }

        private void relKeyWordResults(List<ProductKeyWordList> productKeyWordLists)
        {
            // 네이버 광고 API 실행

            List<ProductKeyWordList> titlelist = productKeyWordLists as List<ProductKeyWordList>;
            NaverApi naverApi = new NaverApi();
            foreach (var v in titlelist)
            {

                string naver = naverApi.NaverAdApi(v.value.Replace(" ", ""));
                JObject obj = JObject.Parse(naver);
                JArray array = null;
                if (obj["keywordList"] != null)
                {
                    array = JArray.Parse(obj["keywordList"].ToString());
                }
                else
                {

                    ReturnToMessage("-------------------------------------------");
                    ReturnToMessage("네이버광고에서 데이터를 불러오지 못했습니다.");
                    ReturnToMessage("현재시간과 컴퓨터시간이 오차가 있는지 확인해주세요.");
                    ReturnToMessage("오차가 있다면 인터넷 시간서버와 동기화 후 다시 시도해주세요.");
                    ReturnToMessage("-------------------------------------------");

                }

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
                }
            }
        }


        private void excelOutFile(string saveFileName, string openFileName)
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;

            object misValue = System.Reflection.Missing.Value;
            
            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            // 엑셀 해더 파일
            xlWorkSheet.Cells[1, 1] = "키워드";
            xlWorkSheet.Cells[1, 2] = "경쟁정도";
            xlWorkSheet.Cells[1, 3] = "노출광고수";



        }
    }
}
