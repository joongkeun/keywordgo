using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
namespace keywordGOGO
{
    class ExcelExport
    {

        public void ExcelCreated(out Excel.Application xlApp, out Excel.Workbook xlWorkBook, out Excel.Worksheet xlWorkSheet, out object misValue)
        {
            misValue = System.Reflection.Missing.Value;
            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            xlWorkSheet.Name = "연관검색어 검색결과";
        }

        public void ExcelHeader(Excel.Worksheet xlWorkSheet)
        {
            xlWorkSheet.Cells[1, 1] = "연관키워드";
            xlWorkSheet.Cells[1, 2] = "월간조회수";
            xlWorkSheet.Cells[1, 3] = "월간평균클릭수";
            xlWorkSheet.Cells[1, 4] = "경쟁상품수";
        }

        public void ExcelData(int r, Excel.Worksheet xlWorkSheet, string RelKeyword, int TotalQcCnt, float TotalCklCnt, int SellPrdQcCnt)
        {
            xlWorkSheet.Cells[r, 1] = RelKeyword; //	연관키워드
            xlWorkSheet.Cells[r, 2] = TotalQcCnt; //	월간조회수
            xlWorkSheet.Cells[r, 3] = TotalCklCnt; //	월간평균클릭수
            xlWorkSheet.Cells[r, 4] = SellPrdQcCnt; //	경쟁상품수
        }


        public void SaveExcel(string saveFileName, object misValue, Excel.Application xlApp, Excel.Workbook xlWorkBook, Excel.Worksheet xlWorkSheet)
        {
            // 파일생성
            xlWorkBook.SaveAs(saveFileName, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            ReleaseObject(xlWorkSheet);
            ReleaseObject(xlWorkBook);
            ReleaseObject(xlApp);
        }



        /// <summary>
        /// 엑셀파일 ReleaseObject 함수
        /// </summary>
        /// <param name="obj"></param>
        public void ReleaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
            }
            finally
            {
                GC.Collect();
            }
        }

        public void ExcelKill()
        {
            // 다 사용한 엑셀 프로세서를 강제 종료한다.
            Process[] ExCel = Process.GetProcessesByName("EXCEL");
            if (ExCel.Count() != 0)
            {
                ExCel[0].Kill();
            }
        }

    }
}
