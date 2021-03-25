using System;
using System.Collections.Generic;
using System.Linq;
using System.Diagnostics;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Newtonsoft.Json.Linq;
using System.Threading;
using System.Windows.Forms;

namespace keywordGOGO
{
    class ExcelTOfile
    {
        // 메인폼 전달 이벤트 선언
        public static event listBoxText ReturnToMessage;
        public static event labelText ReturnToLabel;
        

        public void dataSheet(object keyWord, object SEOkeyWord, object titleKeyWord , string saveFileName)
        {
            List<ExcellOutResult> ExData = new List<ExcellOutResult>();
            List<ExcellOutResult> prdkeword =new List<ExcellOutResult>();
            // 연관 키워드
            List<KeyWordResult> collection = keyWord as List<KeyWordResult>;
            foreach (var r in collection)
            {
                ExData.Add(new ExcellOutResult { RelKeyword = r.RelKeyword, PlAvgDepth = r.PlAvgDepth, Kinds = "연관검색어" });
            }


            // 상품명 키워드
            List<ProductKeyWordList> titlelist = titleKeyWord as List<ProductKeyWordList>;

            foreach (var r in titlelist)
            {
                ExData.Add(new ExcellOutResult { RelKeyword = r.value, PlAvgDepth = "", Kinds = "상품명키워드" });
            }

            // SEO키워드 
            List<KeywordList> seolist = SEOkeyWord as List<KeywordList>;
            foreach (var r in seolist)
            {
                ExData.Add(new ExcellOutResult { RelKeyword = r.Keyword, PlAvgDepth = "", Kinds = "SEO키워드" });
            }


            // 중복 단어의 수를 체크한다.
            var q = ExData.GroupBy(x => x)
           .Select(g => new { Value = g.Key, Count = g.Count() })
           .OrderByDescending(x => x.Count).ToList();
            //중복 키워드를 리스트에 담는다.
            foreach (var temp in q)
            {
                prdkeword.Add(new ExcellOutResult() { RelKeyword = temp.Value.RelKeyword, PlAvgDepth = temp.Value.PlAvgDepth, Kinds = temp.Value.Kinds , Count = temp.Count });
            }

            relKeyWordResults(prdkeword, saveFileName);
        }

        private void relKeyWordResults(List<ExcellOutResult> productKeyWordLists , string saveFileName)
        {
            /*
            // 네이버 광고 API 실행
            List<RelKeyWordResult> relKeyWordResults = new List<RelKeyWordResult>();
            List<ExcellOutResult> titlelist = productKeyWordLists as List<ExcellOutResult>;
            NaverApi naverApi = new NaverApi();
            foreach (var v in titlelist)
            {
                ReturnToLabel(v.RelKeyword.Replace(" ", ""));
                Thread.Sleep(1000);
                string naver = naverApi.NaverAdApi(v.RelKeyword.Replace(" ", ""));


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
                    string plAvgDepth = itemObj["plAvgDepth"].ToString();// 월간노출광고수
                    string compIdx = itemObj["compIdx"].ToString();// 경쟁정도

                    monthlyPcQcCnt = monthlyPcQcCnt.Replace("<", "");
                    monthlyMobileQcCnt = monthlyMobileQcCnt.Replace("<", "");

                    relKeyWordResults.Add(new RelKeyWordResult() { RelKeyword = relKeyword, MonthlyPcQcCnt = monthlyPcQcCnt, MonthlyMobileQcCnt = monthlyMobileQcCnt, MonthlyAvePcClkCnt = monthlyAvePcClkCnt, MonthlyAveMobileClkCnt = monthlyAveMobileClkCnt, MonthlyAvePcCtr = monthlyAvePcCtr, MonthlyAveMobileCtr = monthlyAveMobileCtr, PlAvgDepth = plAvgDepth, CompIdx = compIdx });

                    ReturnToLabel(relKeyword);
                }
            }
            */
            excelOutFile(productKeyWordLists, saveFileName);
        }


        private void excelOutFile(List<ExcellOutResult> relKeyWordResultse, string saveFileName)
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;

            object misValue = System.Reflection.Missing.Value;
            
            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            // 엑셀 해더 파일
            xlWorkSheet.Cells[1, 1] = "연관 키워드";
            xlWorkSheet.Cells[1, 2] = "월간노출 광고수";
            xlWorkSheet.Cells[1, 3] = "키워드 종류";
            //xlWorkSheet.Cells[1, 4] = "중복갯수";


            List<ExcellOutResult> outData = relKeyWordResultse as List<ExcellOutResult>;

            int r = 2;
            foreach (var v in outData)
            {
                ReturnToLabel(v.RelKeyword);

                xlWorkSheet.Cells[r, 1] = v.RelKeyword;
                xlWorkSheet.Cells[r, 2] = v.PlAvgDepth;
                xlWorkSheet.Cells[r, 3] = v.Kinds; 
                //xlWorkSheet.Cells[r, 4] = v.Count; 

                r++;
            }

           ReturnToMessage("엑셀파일을 생성중입니다. 잠시만 기다려 주세요");

            // 파일생성

            xlWorkBook.SaveAs(saveFileName, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();


            ReturnToMessage("보고서 엑셀파일을 생성하였습니다.");
            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);

            // 다 사용한 엑셀 프로세서를 강제 종료한다.
            Process[] ExCel = Process.GetProcessesByName("EXCEL");
            if (ExCel.Count() != 0)
            {
                ExCel[0].Kill();
            }

            
        }

        public void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                ReturnToMessage("프로그램을 Release 하는 도중 오류가 발생하였습니다. : " + ex);
                //ReturnToButton(true);
            }
            finally
            {
                GC.Collect();
            }
        }

    }
}
