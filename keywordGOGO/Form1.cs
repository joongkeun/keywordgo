using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using RestSharp;
using System.Windows.Forms;
using System.Data.SQLite;
using Excel = Microsoft.Office.Interop.Excel;
using System.Threading;
using System.Diagnostics;
using System.IO;

namespace keywordGOGO
{
    //메인폼 전달 델리게이트 선언
    public delegate void listBoxText(string msgText);
    public delegate void labelText(string msgText);

    public partial class Form1 : Form
    {

        delegate void DsetListBox(string data); //리스트박스 델리게이트
        delegate void DsetLabel(string data); //라벨 델리게이트
        delegate void DsetCountLabel(string data, Label label); //라벨 델리게이트
        delegate void DsetApiCntLabel(string data); //라벨 델리게이트
        delegate void DataGrid(List<KeyWordResult> data, DataGridView dataGridView); //데이터그리드 델리게이트
        delegate void DataGrid2(List<ProductKeyWordList> data, DataGridView dataGridView); //데이터그리드 델리게이트
        delegate void DataGrid3(List<string> data, DataGridView dataGridView); //데이터그리드 델리게이트
        delegate void DataGrid4(List<ShopAPIResult> data, DataGridView dataGridView); //데이터그리드 델리게이트
        delegate void DataGrid5(List<KeywordList> data, DataGridView dataGridView); //데이터그리드 델리게이트
        delegate void ButtonEnable(bool data);
        delegate void GridEnable(bool data);

        private SQLiteConnection conn = null;
        private int RefMaxCount = 0;
        private int curRow = -1;

        private GridResultData DataResult = new GridResultData();

        List<KeyWordResult> AllData = new List<KeyWordResult>();

        /// <summary>
        ///  폼 초기화시 
        /// </summary>
        public Form1()
        {
            InitializeComponent();
            this.MaximizeBox = false;
            // 기본 50건을 조회하도록 미리체크
            radioButton1.Checked = true;

            NaverApi.ReturnToMessage += NaverApi_ReturnToMessage;
            NaverApi.ReturnToLabel += NaverApi_ReturnToLabel;
            OutData.ReturnToLabel += OutData_ReturnToLabel; ;
            OutData.ReturnToMessage += OutData_ReturnToMessage; ;
            NaverShoppingCrawler.ReturnToLabel += NaverShoppingCrawler_ReturnToLabel;

        }

        private void OutData_ReturnToMessage(string msgText)
        {
            SetListBox(msgText);
        }

        private void OutData_ReturnToLabel(string msgText)
        {
            SetLabel(msgText);
        }

        /// <summary>
        /// 네이버 쇼핑의 웹페이지를 크롤링 데이터를 라벨에표시합니다.
        /// </summary>
        /// <param name="msgText"></param>
        private void NaverShoppingCrawler_ReturnToLabel(string msgText)
        {
            SetLabel(msgText);
        }

        /// <summary>
        /// 네이버 오픈 API 정보를 라벨에 표시합니다.
        /// </summary>
        /// <param name="msgText"></param>
        private void NaverApi_ReturnToLabel(string msgText)
        {
            SetLabel(msgText);
        }

        /// <summary>
        /// 네이버 오픈 API 정보를 텍스트 박스에 표시합니다.
        /// </summary>
        /// <param name="msgText"></param>
        private void NaverApi_ReturnToMessage(string msgText)
        {
            SetListBox(msgText);
        }

        /// <summary>
        /// 데이터 베이스에 접근해 사용량을 라벨에 표시합니다.
        /// </summary>
        /// <param name="data"></param>
        /// <param name="label"></param>
        public void SetCountLabel(string data, Label label)
        {
            if (label.InvokeRequired)
            {
                DsetCountLabel call = new DsetCountLabel(SetCountLabel);
                this.Invoke(call, data, label);
            }
            else
            {
                label.Text = data;
            }
        }


        public void SetLabel(string data)
        {
            if (label12.InvokeRequired)
            {
                DsetLabel call = new DsetLabel(SetLabel);
                this.Invoke(call, data);
            }
            else
            {

                label12.Text = data;
            }
        }

        public void SetApiCntLabel(string data)
        {
            if (label10.InvokeRequired)
            {
                DsetLabel call = new DsetLabel(SetApiCntLabel);
                this.Invoke(call, data);
            }
            else
            {
                label10.Text = data;
            }
        }

        public void SetListBox(string data)
        {
            if (listBox1.InvokeRequired)
            {

                DsetListBox call = new DsetListBox(SetListBox);
                this.Invoke(call, data);
            }
            else
            {
                listBox1.Items.Add(data);
                listBox1.SelectedIndex = listBox1.Items.Count - 1;
            }
        }

        private void SetButton(bool data)
        {
            if (saveBtn.InvokeRequired)

            {

                ButtonEnable call = new ButtonEnable(SetButton);
                this.Invoke(call, data);

            }

            else
            {
                saveBtn.Enabled = data;
            }
        }


        private void SetGrid(bool data)
        {
            if (dataGridView6.InvokeRequired)

            {

                GridEnable call = new GridEnable(SetGrid);
                this.Invoke(call, data);

            }

            else
            {
                dataGridView6.Enabled = data;
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            try
            {
                string strConn = @"Data Source=" + Application.StartupPath + "\\apiQc.db";
                conn = new SQLiteConnection(strConn);
                conn.Open();
                SetListBox("+++++++++++++++++++++++++++++++++++++");
                SetListBox("데이터베이스 연결성공");
                SetListBox("+++++++++++++++++++++++++++++++++++++");

                // api 사용량 
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

                SetApiCntLabel(count);
                int rCount = 0;
                if (string.IsNullOrEmpty(count))
                {

                    rCount = 0;
                }
                else
                {
                    rCount = Convert.ToInt32(count);
                }

                if (rCount >= 25000)
                {
                    SetButton(false);
                    SetListBox("네이버 Open Api 일일 사용량을 초과하여 더이상 조회를 할 수 없습니다.");
                }

            }
            catch
            {
                SetButton(false);
                SetListBox("데이터베이스 연결에 실패했습니다.");
            }


            iniUtil ini = new iniUtil(Application.StartupPath + "\\config.ini");

   
            string apiKey = ini.GetIniValue("ADAPI", "apiKey");
            string secretKey = ini.GetIniValue("ADAPI", "secretKey");
            string managerCustomerId = ini.GetIniValue("ADAPI", "managerCustomerId");
            string ClientId  =ini.GetIniValue("OPENAPI", "ClientId"); // 클라이언트 아이디
            string ClientSecret = ini.GetIniValue("OPENAPI", "ClientSecret");       // 클라이언트 시크릿

            if(string.IsNullOrEmpty(apiKey) || string.IsNullOrEmpty(secretKey) || string.IsNullOrEmpty(managerCustomerId) || string.IsNullOrEmpty(ClientId) || string.IsNullOrEmpty(ClientSecret))
            {
                MessageBox.Show("config.ini 파일을 확인 하십시오.", "경고", MessageBoxButtons.OK);
                Process.Start("notepad.exe", "config.ini");
                Process.Start("iexplore.exe", "https://vitdeul.tistory.com/8");
                Application.ExitThread();
                Environment.Exit(0);

            }

            webBrowser1.Navigate("https://vitdeul.tistory.com/8");
        }

        private void saveBtn_Click(object sender, EventArgs e)
        {
            listBox1.Items.Clear();
            if (radioButton1.Checked == true) RefMaxCount = 50;
            if (radioButton2.Checked == true) RefMaxCount = 100;
            if (radioButton3.Checked == true) RefMaxCount = 200;
            if (radioButton4.Checked == true) RefMaxCount = 300;
            Thread t1 = new Thread(new ThreadStart(DataReturn));
            t1.Start();
        }


        public void SetDataGrid(object msgData, DataGridView dataGridView)
        {
            if (dataGridView.InvokeRequired)
            {
                DataGrid call = new DataGrid(SetDataGrid);
                this.Invoke(call, msgData, dataGridView);
            }
            else
            {
                dataGridView.Rows.Clear();
                List<KeyWordResult> collection = msgData as List<KeyWordResult>;
                dataGridView.ColumnHeadersVisible = true;
                dataGridView.RowHeadersVisible = false;
                DataGridViewCellStyle columnHeaderStyle = new DataGridViewCellStyle();
                columnHeaderStyle.Font = new Font("Veradna", 10, FontStyle.Bold);
                columnHeaderStyle.BackColor = Color.Beige;
                dataGridView.ColumnHeadersDefaultCellStyle = columnHeaderStyle;
                dataGridView.ColumnCount = 10;
                dataGridView.Columns[0].HeaderCell.Value = "키워드";
                dataGridView.Columns[1].HeaderCell.Value = "월간 PC 검색수";
                dataGridView.Columns[2].HeaderCell.Value = "월간 모바일 검색수";
                dataGridView.Columns[3].HeaderCell.Value = "월간  PC 클릭수";
                dataGridView.Columns[4].HeaderCell.Value = "월간 모바일 클릭수";
                dataGridView.Columns[5].HeaderCell.Value = "월간 PC 클릭률";
                dataGridView.Columns[6].HeaderCell.Value = "월간 모바일 클릭률";
                dataGridView.Columns[7].HeaderCell.Value = "경쟁정도";
                dataGridView.Columns[8].HeaderCell.Value = "월간 노출 광고수";
                dataGridView.Columns[9].HeaderCell.Value = "검색 상품수";

                foreach (var r in collection)
                {
                    dataGridView.Rows.Add(r.RelKeyword, r.MonthlyPcQcCnt, r.MonthlyMobileQcCnt, r.MonthlyAvePcClkCnt, r.MonthlyAveMobileClkCnt, r.MonthlyAvePcCtr, r.MonthlyAveMobileCtr, r.PlAvgDepth, r.CompIdx, r.SellPrdQcCnt);
                }
            }
        }

        public void DataReturn()
        {
            SetButton(false);
            SetListBox("데이터를 조회하기 위해 초기화 중입니다.");
            OutData outData = new OutData();
            DataResult = outData.GridDataSet(keywordTbox.Text, RefMaxCount, conn);

            SetDataGrid(DataResult.AdRefGrid, dataGridView6); //전체 연관 검색어 리스트

            // api 사용량 
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

            SetApiCntLabel(count);
            SetButton(true);
        }

        /// <summary>
        /// 그리드 데이터를 선택하면 데이터를 더 불러온다.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dataGridView6_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView6.CurrentRow.Index != curRow)
            {
                if (dataGridView6.CurrentRow.Index != DataResult.AdRefGrid.Count)
                {
                    curRow = dataGridView6.CurrentRow.Index;
                    string data = dataGridView6.CurrentRow.Cells[0].Value.ToString();

                    Console.WriteLine(curRow);
                    Console.WriteLine(data);

                    SetDataGridClear();

                    Thread t2 = new Thread(() => SubDataReturn(data));
                    t2.Start();


                }
            }
        }

        public void SubDataReturn(string data)
        {

            SetGrid(false);


            GridResultData SubDataResult = new GridResultData();

            List<ProductKeyWordList> ProductWordList = new List<ProductKeyWordList>();
            List<string> MallList = new List<string>();
            List<string> PrdKeyWord = new List<string>();
            OutData outData = new OutData();

            var shopResult = (from a in DataResult.AdRefGrid where a.RelKeyword == data select a.ShopResult);
           
            foreach (var result in shopResult)
            {
                foreach (var a in result)
                {
                    foreach (var keyword in a.TitleKeywordList)
                    {
                        SetLabel(keyword);

                        bool BoolValue = true;
                        string[] tempData = new string[3] { " ", "/", "-" };
                        foreach (var d in tempData)
                        {
                            if (keyword == d)
                            {
                                BoolValue = false;
                            }
                            else if (keyword.Length == 0)
                            {
                                BoolValue = false;
                            }
                        }

                        if (BoolValue == true)
                        {
                            PrdKeyWord.Add(keyword);
                        }
                    }

                    //q = q.Take(100).ToList();
                    if (a.MallName != "네이버")
                    {
                        MallList.Add(a.MallName);
                    }
                }
            }

            // 중복 단어의 수를 체크한다.
            var q = PrdKeyWord.GroupBy(x => x)
           .Select(g => new { Value = g.Key, Count = g.Count() })
           .OrderByDescending(x => x.Count).ToList();
            //중복 키워드를 리스트에 담는다.
            foreach (var temp in q)
            {
                ProductWordList.Add(new ProductKeyWordList() { value = temp.Value, count = temp.Count });
            }

  
            SubDataResult = outData.SubGridDataSet(data);

            // 상품정보
            foreach (var result in shopResult)
            {
                SetDataGrid4(result, dataGridView7);
            }

            // 상품명 분석
            SetDataGrid2(ProductWordList, dataGridView4);

            //연관검색어
            SetDataGrid5(SubDataResult.ShoppingRefGrid, dataGridView5);

            //연관검색어
            SetDataGrid5(SubDataResult.ShopWebDataResult.OutTagList, dataGridView1);
            
            SetGrid(true);

        }

        public void SetDataGrid2(object msgData, DataGridView dataGridView)
        {
            if (dataGridView.InvokeRequired)
            {
                DataGrid2 call = new DataGrid2(SetDataGrid2);
                this.Invoke(call, msgData, dataGridView);
            }
            else
            {
                dataGridView.Rows.Clear();

                List<ProductKeyWordList> collection = msgData as List<ProductKeyWordList>;

                dataGridView.Rows.Clear();

                dataGridView.ColumnHeadersVisible = true;
                dataGridView.RowHeadersVisible = false;
                DataGridViewCellStyle columnHeaderStyle = new DataGridViewCellStyle();
                columnHeaderStyle.Font = new Font("Veradna", 10, FontStyle.Bold);
                columnHeaderStyle.BackColor = Color.Beige;
                dataGridView.ColumnHeadersDefaultCellStyle = columnHeaderStyle;
                dataGridView.ColumnCount = 2;
                dataGridView.Columns[0].HeaderCell.Value = "키워드";
                dataGridView.Columns[1].HeaderCell.Value = "사용량";

                foreach (var r in collection)
                {
                    dataGridView.Rows.Add(r.value, r.count);
                }
            }
        }

        public void SetDataGrid4(object msgData, DataGridView dataGridView)
        {
            if (dataGridView.InvokeRequired)
            {
                DataGrid4 call = new DataGrid4(SetDataGrid4);
                this.Invoke(call, msgData, dataGridView);
            }
            else
            {
                dataGridView.Rows.Clear();

                List<ShopAPIResult> collection = msgData as List<ShopAPIResult>;

                dataGridView.Rows.Clear();

                dataGridView.ColumnHeadersVisible = true;
                dataGridView.RowHeadersVisible = false;
                DataGridViewCellStyle columnHeaderStyle = new DataGridViewCellStyle();
                columnHeaderStyle.Font = new Font("Veradna", 10, FontStyle.Bold);
                columnHeaderStyle.BackColor = Color.Beige;
                dataGridView.ColumnHeadersDefaultCellStyle = columnHeaderStyle;
                dataGridView.ColumnCount = 6;
                dataGridView.Columns[0].HeaderCell.Value = "상품명";
                dataGridView.Columns[1].HeaderCell.Value = "몰이름";
                dataGridView.Columns[2].HeaderCell.Value = "최저가";
                dataGridView.Columns[3].HeaderCell.Value = "최고가";
                dataGridView.Columns[4].HeaderCell.Value = "상품ID";
                dataGridView.Columns[5].HeaderCell.Value = "상품주소";

                foreach (var a in collection)
                {
                    string mall = string.Empty;
                    if (a.MallName == "네이버")
                    {
                        mall = "가격비교";
                    }
                    else
                    {
                        mall = a.MallName;
                    }

                    dataGridView.Rows.Add(a.Title, mall, a.Lprice, a.Hprice, a.ProductId, a.Link);
                }
            }
        }

        public void SetDataGrid5(object msgData, DataGridView dataGridView)
        {
            if (dataGridView.InvokeRequired)
            {
                DataGrid5 call = new DataGrid5(SetDataGrid5);
                this.Invoke(call, msgData, dataGridView);
            }
            else
            {
                dataGridView.Rows.Clear();

                List<KeywordList> collection = msgData as List<KeywordList>;

                dataGridView.Rows.Clear();

                dataGridView.ColumnHeadersVisible = true;
                dataGridView.RowHeadersVisible = false;
                DataGridViewCellStyle columnHeaderStyle = new DataGridViewCellStyle();
                columnHeaderStyle.Font = new Font("Veradna", 10, FontStyle.Bold);
                columnHeaderStyle.BackColor = Color.Beige;
                dataGridView.ColumnHeadersDefaultCellStyle = columnHeaderStyle;
                dataGridView.ColumnCount = 1;
                dataGridView.Columns[0].HeaderCell.Value = "키워드";

                int count = 1;
                foreach (var r in collection)
                {
                    if (r.Kind == "S" || r.Kind == "T")
                    {
                        dataGridView.Rows.Add(r.Keyword);
                    }

                }
            }
        }

       public void SetDataGridClear()
       {
            dataGridView1.Rows.Clear();
            dataGridView4.Rows.Clear();
            dataGridView5.Rows.Clear();
            dataGridView7.Rows.Clear();
        }
    }
}
