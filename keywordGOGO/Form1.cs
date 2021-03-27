using Microsoft.Win32;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Security;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace keywordGOGO
{
    //메인폼 전달 델리게이트 선언
    public delegate void listBoxText(string msgText);
    public delegate void labelText(string msgText);
   
    public partial class Form1 : Form
    {
        string version = "1.10.1";

        string reportSaveFileName = string.Empty; // 보고서 파일 생성

        delegate void DsetListBox(string data); //리스트박스 델리게이트
        delegate void DsetLabel(string data); //라벨 델리게이트
        delegate void DsetCountLabel(string data, Label label); //라벨 델리게이트
        delegate void DsetApiCntLabel(string data); //라벨 델리게이트
        delegate void DataGrid(List<KeyWordResult> data, DataGridView dataGridView); //데이터그리드 델리게이트
        delegate void DataGrid2(List<ProductKeyWordList> data, DataGridView dataGridView); //데이터그리드 델리게이트
        delegate void DataGrid3(List<string> data, DataGridView dataGridView); //데이터그리드 델리게이트
        delegate void DataGrid4(List<ShopAPIResult> data, DataGridView dataGridView); //데이터그리드 델리게이트
        delegate void DataGrid5(List<KeywordList> data, DataGridView dataGridView); //데이터그리드 델리게이트

        // 순위
        delegate void DataGrid6(List<RankingList> data, DataGridView dataGridView); //데이터그리드 델리게이트
        delegate void DataGrid7(List<RankingList> data, DataGridView dataGridView); //데이터그리드 델리게이트
        delegate void DataGrid9(List<RankingList> data, DataGridView dataGridView); //데이터그리드 델리게이트

        // 인스타
        delegate void DataGrid8(List<InstagramTagWordList> data, DataGridView dataGridView); //데이터그리드 델리게이트


        delegate void ButtonEnable(bool data);
        delegate void GridEnable(bool data);

        private SQLiteConnection conn = null;
        private int RefMaxCount = 0;
        private int curRow = -1;

        private GridResultData DataResult = new GridResultData();

        List<KeyWordResult> AllData = new List<KeyWordResult>();

        private System.Windows.Forms.Timer timer = new System.Windows.Forms.Timer();

        // 웹브라우져 레지스트리 변경
        private const string InternetExplorerRootKey = @"Software\Microsoft\Internet Explorer";
        private const string BrowserEmulationKey = InternetExplorerRootKey + @"\Main\FeatureControl\FEATURE_BROWSER_EMULATION";



        /// <summary>
        ///  폼 초기화시 
        /// </summary>
        public Form1()
        {
            InitializeComponent();
            //this.MaximizeBox = false;
            // 기본 50건을 조회하도록 미리체크
            radioButton1.Checked = true;

            NaverApi.ReturnToMessage += NaverApi_ReturnToMessage;
            NaverApi.ReturnToLabel += NaverApi_ReturnToLabel;
            OutData.ReturnToLabel += OutData_ReturnToLabel; ;
            OutData.ReturnToMessage += OutData_ReturnToMessage; ;
            NaverShoppingCrawler.ReturnToLabel += NaverShoppingCrawler_ReturnToLabel;
            NaverShoppingCrawler.ReturnToMessage += NaverShoppingCrawler_ReturnToMessage;
            bizranking.ReturnToLabel += Bizranking_ReturnToLabel;
            bizranking.ReturnToMessage += Bizranking_ReturnToMessage;
            ExcelTOfile.ReturnToLabel += ExcelTOfile_ReturnToLabel;
            ExcelTOfile.ReturnToMessage += ExcelTOfile_ReturnToMessage;
            timer.Tick += Timer_Tick;
        }

        private void ExcelTOfile_ReturnToMessage(string msgText)
        {
            SetListBox(msgText);
        }

        private void ExcelTOfile_ReturnToLabel(string msgText)
        {
            SetLabel2(msgText);
        }

        private void NaverShoppingCrawler_ReturnToMessage(string msgText)
        {
            SetListBox(msgText);
        }

        private void Bizranking_ReturnToMessage(string msgText)
        {
            SetListBox2(msgText);
        }

        private void Bizranking_ReturnToLabel(string msgText)
        {
            SetLabel2(msgText);
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

        public void SetLabel2(string data)
        {
            if (label12.InvokeRequired)
            {
                DsetLabel call = new DsetLabel(SetLabel2);
                this.Invoke(call, data);
            }
            else
            {

                label15.Text = data;
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

        public void SetListBox2(string data)
        {
            if (listBox1.InvokeRequired)
            {

                DsetListBox call = new DsetListBox(SetListBox2);
                this.Invoke(call, data);
            }
            else
            {
                listBox2.Items.Add(data);
                listBox2.SelectedIndex = listBox2.Items.Count - 1;
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


        private void SetButton2(bool data)
        {
            if (button2.InvokeRequired)

            {

                ButtonEnable call = new ButtonEnable(SetButton2);
                this.Invoke(call, data);

            }

            else
            {
                button2.Enabled = data;
            }
        }


        private void SetInstaButton(bool data)
        {
            if (instarTagBtn.InvokeRequired)

            {

                ButtonEnable call = new ButtonEnable(SetInstaButton);
                this.Invoke(call, data);

            }

            else
            {
                instarTagBtn.Enabled = data;
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


        public string VersionChk()
        {
            try
            {
                string url = "http://193.123.251.70/api/version/read_version.php"; // 결과가 JSON 포맷
                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);
                HttpWebResponse response = (HttpWebResponse)request.GetResponse();

                Stream stream = response.GetResponseStream();
                StreamReader reader = new StreamReader(stream, Encoding.UTF8);
                string text = reader.ReadToEnd();

                JObject obj = JObject.Parse(text);
                string version = obj["version"].ToString();

                Console.WriteLine(version);
                return version;
            }
            catch
            {

                return string.Empty;
            }

        }


        public string NoticeChk()
        {
            try
            {
                string url = "http://193.123.251.70/api/version/read_version.php"; // 결과가 JSON 포맷
                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);
                HttpWebResponse response = (HttpWebResponse)request.GetResponse();

                Stream stream = response.GetResponseStream();
                StreamReader reader = new StreamReader(stream, Encoding.UTF8);
                string text = reader.ReadToEnd();

                JObject obj = JObject.Parse(text);
                string notice = obj["message"].ToString();

                Console.WriteLine(notice);
                return notice;
            }
            catch
            {

                return string.Empty;
            }

        }


        private bool sqlliteDBChk()
        {
            string DbFile = "apiQc.db";
            string strConn = @"Data Source=" + Application.StartupPath + "\\apiQc.db";
            try
            {
                // 파일생성
                SQLiteConnection.CreateFile(DbFile);
                // 테이블 생성 코드
                SQLiteConnection sqliteConn = new SQLiteConnection(strConn);
                sqliteConn.Open();

                string strsql = "create table if not exists apicount (date text, count int,primary key(date))";

                SQLiteCommand cmd = new SQLiteCommand(strsql, sqliteConn);
                cmd.ExecuteNonQuery();
                sqliteConn.Close();
                return true;
            }
            catch
            {
                return false;
            }

        }


        private void Form1_Load(object sender, EventArgs e)
        {
            string notice = "";
            try
            {
                notice = NoticeChk();
            }
            catch
            {
                notice = "";
            }

            label20.Text = notice;

            try
            {

                this.Text = "키워드고고(v" + version + ")";

                // 버전확인
                string versionchk = VersionChk();
                if (versionchk != version)
                {
                    MessageBox.Show("새버전이 있습니다. 도움말을 참고하여 업데이트 하십시오.", "경고", MessageBoxButtons.OK);
                }




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
                //SetButton(false);
                sqlliteDBChk();
                SetListBox("데이터베이스 연결에 실패하여 새로 생성하였습니다.");
            }


            iniUtil ini = new iniUtil(Application.StartupPath + "\\config.ini");


            string apiKey = ini.GetIniValue("ADAPI", "apiKey");
            string secretKey = ini.GetIniValue("ADAPI", "secretKey");
            string managerCustomerId = ini.GetIniValue("ADAPI", "managerCustomerId");
            string ClientId = ini.GetIniValue("OPENAPI", "ClientId"); // 클라이언트 아이디
            string ClientSecret = ini.GetIniValue("OPENAPI", "ClientSecret");       // 클라이언트 시크릿

            if (string.IsNullOrEmpty(apiKey) || string.IsNullOrEmpty(secretKey) || string.IsNullOrEmpty(managerCustomerId) || string.IsNullOrEmpty(ClientId) || string.IsNullOrEmpty(ClientSecret))
            {
                MessageBox.Show("config.ini 파일을 확인 하십시오.", "경고", MessageBoxButtons.OK);
                Process.Start("notepad.exe", "config.ini");
                Process.Start("iexplore.exe", "https://vitdeul.tistory.com/8");
                Application.ExitThread();
                Environment.Exit(0);

            }


            SetBrowserEmulationVersion();


            var additionalHeaders = "User-Agent:Mozilla/5.0 (Windows Phone 10.0; Android 6.0.1; " +
                                        "Microsoft; Lumia 950 XL Dual SIM) AppleWebKit/537.36 (KHTML, like Gecko) " +
                                        "Chrome/52.0.2743.116 Mobile Safari/537.36 Edge/15.15063\r\n";

            this.webBrowser3.Navigate("https://m.shopping.naver.com/home/m/index.nhn", null, null, additionalHeaders);


            webBrowser7.Navigate("https://vitdeul.tistory.com/8");
            webBrowser2.Navigate("https://datalab.naver.com/shoppingInsight/sCategory.naver");

            webBrowser4.Navigate("https://shopping.naver.com/home/p/index.nhn");
            //webBrowser5.Navigate("https://blackkiwi.net");

            //webBrowser6.Navigate("https://www.instagram.com/");
            webBrowser8.Navigate("https://analytics.naver.com/");
            checkBox2.Checked = true;
        }

        public enum BrowserEmulationVersion

        {

            Default = 0,

            Version7 = 7000,

            Version8 = 8000,

            Version8Standards = 8888,

            Version9 = 9000,

            Version9Standards = 9999,

            Version10 = 10000,

            Version10Standards = 10001,

            Version11 = 11000,

            Version11Edge = 11001

        }
        public static int GetInternetExplorerMajorVersion()

        {

            int result;



            result = 0;



            try

            {

                RegistryKey key;



                key = Registry.LocalMachine.OpenSubKey(InternetExplorerRootKey);



                if (key != null)

                {

                    object value;



                    value = key.GetValue("svcVersion", null) ?? key.GetValue("Version", null);



                    if (value != null)

                    {

                        string version;

                        int separator;



                        version = value.ToString();

                        separator = version.IndexOf('.');

                        if (separator != -1)

                        {

                            int.TryParse(version.Substring(0, separator), out result);

                        }

                    }
                }
            }

            catch (SecurityException)

            {

                // The user does not have the permissions required to read from the registry key.

            }

            catch (UnauthorizedAccessException)

            {

                // The user does not have the necessary registry rights.

            }



            return result;

        }

        public static BrowserEmulationVersion GetBrowserEmulationVersion()

        {

            BrowserEmulationVersion result;



            result = BrowserEmulationVersion.Default;



            try

            {

                RegistryKey key;



                key = Registry.CurrentUser.OpenSubKey(BrowserEmulationKey, true);

                if (key != null)

                {

                    string programName;

                    object value;



                    programName = Path.GetFileName(Environment.GetCommandLineArgs()[0]);

                    value = key.GetValue(programName, null);



                    if (value != null)

                    {

                        result = (BrowserEmulationVersion)Convert.ToInt32(value);

                    }

                }

            }

            catch (SecurityException)

            {

                // The user does not have the permissions required to read from the registry key.

            }

            catch (UnauthorizedAccessException)

            {

                // The user does not have the necessary registry rights.

            }



            return result;

        }



        public static bool IsBrowserEmulationSet()

        {

            return GetBrowserEmulationVersion() != BrowserEmulationVersion.Default;

        }

        public static bool SetBrowserEmulationVersion(BrowserEmulationVersion browserEmulationVersion)

        {

            bool result;



            result = false;



            try

            {

                RegistryKey key;



                key = Registry.CurrentUser.OpenSubKey(BrowserEmulationKey, true);



                if (key != null)

                {

                    string programName;



                    programName = Path.GetFileName(Environment.GetCommandLineArgs()[0]);



                    if (browserEmulationVersion != BrowserEmulationVersion.Default)

                    {

                        // if it's a valid value, update or create the value

                        key.SetValue(programName, (int)browserEmulationVersion, RegistryValueKind.DWord);

                    }

                    else

                    {

                        // otherwise, remove the existing value

                        key.DeleteValue(programName, false);

                    }



                    result = true;

                }

            }

            catch (SecurityException)

            {

                // The user does not have the permissions required to read from the registry key.

            }

            catch (UnauthorizedAccessException)

            {

                // The user does not have the necessary registry rights.

            }



            return result;

        }



        public static bool SetBrowserEmulationVersion()

        {

            int ieVersion;

            BrowserEmulationVersion emulationCode;



            ieVersion = GetInternetExplorerMajorVersion();



            if (ieVersion >= 11)

            {

                emulationCode = BrowserEmulationVersion.Version11;

            }

            else

            {

                switch (ieVersion)

                {

                    case 10:

                        emulationCode = BrowserEmulationVersion.Version10;

                        break;

                    case 9:

                        emulationCode = BrowserEmulationVersion.Version9;

                        break;

                    case 8:

                        emulationCode = BrowserEmulationVersion.Version8;

                        break;

                    default:

                        emulationCode = BrowserEmulationVersion.Version7;

                        break;

                }

            }



            return SetBrowserEmulationVersion(emulationCode);

        }



        private void saveBtn_Click(object sender, EventArgs e)
        {
            listBox1.Items.Clear();
           
            SetDataGridClear();


            if (keywordTbox.Text.Length < 1)
            {
                MessageBox.Show("키워드를 넣어주세요!", "경고", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (checkBox1.Checked == true)
            {


                using (SaveFileDialog dlg = new SaveFileDialog())
                {
                    dlg.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
                    dlg.Filter = "모든파일 (*.*)|*.*|모든파일 (*.*)|*.*";
                    dlg.FilterIndex = 1;
                    dlg.RestoreDirectory = true;

                    if (dlg.ShowDialog() != DialogResult.OK)
                    {
                        return;
                    }

                    reportSaveFileName = dlg.FileName;
                    SetListBox(reportSaveFileName);
                }


            }

            if (radioButton1.Checked == true) RefMaxCount = 50;
            if (radioButton2.Checked == true) RefMaxCount = 100;
            if (radioButton3.Checked == true) RefMaxCount = 300;
            if (radioButton4.Checked == true) RefMaxCount = 1000;
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
                dataGridView.ColumnCount = 11;
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
                dataGridView.Columns[10].HeaderCell.Value = "상품수대비 키워드 경쟁강도";

                foreach (var r in collection)
                {
                    int MonthlyPcQcCnt = Convert.ToInt32(r.MonthlyPcQcCnt);
                    int MonthlyMobileQcCnt = Convert.ToInt32(r.MonthlyMobileQcCnt);
                    int SellPrdQcCnt = Convert.ToInt32(r.SellPrdQcCnt);
                    double MonthlyAvePcClkCnt = Convert.ToDouble(r.MonthlyAvePcClkCnt);
                    double MonthlyAveMobileClkCnt = Convert.ToDouble(r.MonthlyAveMobileClkCnt);
                    double MonthlyAvePcCtr = Convert.ToDouble(r.MonthlyAvePcCtr);
                    double MonthlyAveMobileCtr = Convert.ToDouble(r.MonthlyAveMobileCtr);
                    double sellprdQcCompldx = 0;
                    int plAvgDepth = Convert.ToInt32(r.PlAvgDepth);
                    if (SellPrdQcCnt > 0) {
                         sellprdQcCompldx = Convert.ToDouble((MonthlyPcQcCnt + MonthlyMobileQcCnt) * 100 / SellPrdQcCnt);
                    }
                    dataGridView.Rows.Add(r.RelKeyword, string.Format("{0:#,0}", MonthlyPcQcCnt), string.Format("{0:#,0}", MonthlyMobileQcCnt), MonthlyAvePcClkCnt, MonthlyAveMobileClkCnt, MonthlyAvePcCtr, MonthlyAveMobileCtr, plAvgDepth, r.CompIdx , string.Format("{0:#,0}", SellPrdQcCnt), string.Format("{0:#,0}", sellprdQcCompldx));
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
            // SEO태그 검색유무
            bool tagYn = checkBox2.Checked;
            SubDataReturn(keywordTbox.Text, tagYn);

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

        private void dataGridView6_Click(object sender, EventArgs e)
        {
           
        }

        private void dataGridView6_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridView6.CurrentCell == null || dataGridView6.CurrentCell.Value == null || e.RowIndex == -1)
            {
                return;
            }

            if (dataGridView6.CurrentCell.ColumnIndex.Equals(0))
            {
                string data = dataGridView6.CurrentCell.Value.ToString();
                // SEO태그 검색유무
                bool tagYn = checkBox2.Checked;

                if (dataGridView6.CurrentRow.Index != curRow)
                {
                    if (dataGridView6.CurrentRow.Index >= 0)
                    {
                        curRow = dataGridView6.CurrentRow.Index;
                        if (dataGridView6.CurrentRow.Cells[0].Value != null)
                        {
                            data = dataGridView6.CurrentRow.Cells[0].Value.ToString();
                        }
                        else
                        {
                            return;
                        }


                        Console.WriteLine(curRow);
                        Console.WriteLine(data);

                        SetDataGridClear();

                        Thread t2 = new Thread(() => SubDataReturn(data, tagYn));
                        t2.Start();


                    }
                }

            }
        }

        /// <summary>
        /// 그리드 데이터를 선택하면 데이터를 더 불러온다.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dataGridView6_SelectionChanged(object sender, EventArgs e)
        {
           
        }


        public void SubDataReturn(string data, bool tagYn)
        {

            SetGrid(false);
            SetButton(false);

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

            
            SubDataResult = outData.SubGridDataSet(data, tagYn);

            // 상품정보
            foreach (var result in shopResult)
            {
                SetDataGrid4(result, dataGridView7);
            }

            // 상품명 분석
            SetDataGrid2(ProductWordList, dataGridView4);

            //연관검색어
            SetDataGrid5(SubDataResult.ShoppingRefGrid, dataGridView5);

            //SEO 태그
            SetDataGrid5(SubDataResult.ShopWebDataResult.OutTagList, dataGridView1);

            ExcelTOfile excelTOfile = new ExcelTOfile();
            excelTOfile.dataSheet(DataResult.AdRefGrid, SubDataResult.ShopWebDataResult.OutTagList, ProductWordList, reportSaveFileName);

            SetGrid(true);
            SetButton(true);

        }


        public void SetInstaDataGrid(object msgData, DataGridView dataGridView)
        {
            if (dataGridView.InvokeRequired)
            {
                DataGrid8 call = new DataGrid8(SetInstaDataGrid);
                this.Invoke(call, msgData, dataGridView);
            }
            else
            {
                dataGridView.Rows.Clear();

                List<InstagramTagWordList> collection = msgData as List<InstagramTagWordList>;

                dataGridView.Rows.Clear();

                dataGridView.ColumnHeadersVisible = true;
                dataGridView.RowHeadersVisible = false;
                DataGridViewCellStyle columnHeaderStyle = new DataGridViewCellStyle();
                columnHeaderStyle.Font = new Font("Veradna", 10, FontStyle.Bold);
                columnHeaderStyle.BackColor = Color.Beige;
                dataGridView.ColumnHeadersDefaultCellStyle = columnHeaderStyle;
                dataGridView.ColumnCount = 2;
                dataGridView.Columns[0].HeaderCell.Value = "해시태그";
                dataGridView.Columns[1].HeaderCell.Value = "사용량";

                foreach (var r in collection)
                {

                    int count = Convert.ToInt32(r.count);
                    if (r.value.Length != 0)
                    {
                        dataGridView.Rows.Add(r.value, count);
                    }
                }
            }
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

                    int count = Convert.ToInt32(r.count);

                    dataGridView.Rows.Add(r.value, count);
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
                    int Lprice = 0;
                    int Hprice = 0;

                    if (a.MallName == "네이버")
                    {
                        mall = "가격비교";
                    }
                    else
                    {
                        mall = a.MallName;
                    }

                    if (a.Lprice.Equals(""))
                    {
                        Lprice = 0;
                    }
                    else
                    {
                        Lprice = Convert.ToInt32(a.Lprice);
                    }

                    if (a.Hprice.Equals(""))
                    {
                        Hprice = 0;
                    }
                    else
                    {
                        Hprice = Convert.ToInt32(a.Hprice);
                    }




                    dataGridView.Rows.Add(a.Title, mall, string.Format("{0:#,0}", Lprice), string.Format("{0:#,0}", Hprice), a.ProductId, a.Link);
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


        public void SetDataAdRankGrid(object msgData, DataGridView dataGridView)
        {
            if (dataGridView.InvokeRequired)
            {
                DataGrid6 call = new DataGrid6(SetDataAdRankGrid);
                this.Invoke(call, msgData, dataGridView);
            }
            else
            {
                dataGridView.Rows.Clear();

                List<RankingList> collection = msgData as List<RankingList>;

                dataGridView.Rows.Clear();

                dataGridView.ColumnHeadersVisible = true;
                dataGridView.RowHeadersVisible = false;
                DataGridViewCellStyle columnHeaderStyle = new DataGridViewCellStyle();
                columnHeaderStyle.Font = new Font("Veradna", 10, FontStyle.Bold);
                columnHeaderStyle.BackColor = Color.Beige;
                dataGridView.ColumnHeadersDefaultCellStyle = columnHeaderStyle;
                dataGridView.ColumnCount = 11;
                dataGridView.Columns[0].HeaderCell.Value = "현재순위";
                dataGridView.Columns[1].HeaderCell.Value = "과거순위";
                dataGridView.Columns[2].HeaderCell.Value = "현재광고비";
                dataGridView.Columns[3].HeaderCell.Value = "과거광고비";
                dataGridView.Columns[4].HeaderCell.Value = "상품명";
                dataGridView.Columns[5].HeaderCell.Value = "상품번호";
                dataGridView.Columns[6].HeaderCell.Value = "카테고리";
                dataGridView.Columns[7].HeaderCell.Value = "키워드유사성";
                dataGridView.Columns[8].HeaderCell.Value = "키워드관련성";
                dataGridView.Columns[9].HeaderCell.Value = "관련성랭킹";
                dataGridView.Columns[10].HeaderCell.Value = "상품주소";

                int count = 1;
                foreach (var r in collection)
                {

                    dataGridView.Rows.Add(r.pageNo + "페이지 " + r.rank + "위",
                    r.oldPageNo + "페이지 " + r.oldrank + "위",
                    r.adprice,
                    r.oldadprice,
                    r.productName,
                    r.productNo,
                    r.categoryName,
                    r.similarity + "% ",
                    r.relevance + "% ",
                    r.hitRank,
                    r.productUrl);

                }
            }
        }

        public void SetDataRankGrid(object msgData, DataGridView dataGridView)
        {
            if (dataGridView.InvokeRequired)
            {
                DataGrid7 call = new DataGrid7(SetDataRankGrid);
                this.Invoke(call, msgData, dataGridView);
            }
            else
            {
                dataGridView.Rows.Clear();

                List<RankingList> collection = msgData as List<RankingList>;

                dataGridView.Rows.Clear();

                dataGridView.ColumnHeadersVisible = true;
                dataGridView.RowHeadersVisible = false;
                DataGridViewCellStyle columnHeaderStyle = new DataGridViewCellStyle();
                columnHeaderStyle.Font = new Font("Veradna", 10, FontStyle.Bold);
                columnHeaderStyle.BackColor = Color.Beige;
                dataGridView.ColumnHeadersDefaultCellStyle = columnHeaderStyle;
                dataGridView.ColumnCount = 9;
                dataGridView.Columns[0].HeaderCell.Value = "현재순위";
                dataGridView.Columns[1].HeaderCell.Value = "과거순위";
                //dataGridView.Columns[2].HeaderCell.Value = "현재광고비";
                //dataGridView.Columns[3].HeaderCell.Value = "과거광고비";
                dataGridView.Columns[2].HeaderCell.Value = "상품명";
                dataGridView.Columns[3].HeaderCell.Value = "상품번호";
                dataGridView.Columns[4].HeaderCell.Value = "카테고리";
                dataGridView.Columns[5].HeaderCell.Value = "키워드유사성";
                dataGridView.Columns[6].HeaderCell.Value = "키워드관련성";
                dataGridView.Columns[7].HeaderCell.Value = "관련성랭킹";
                dataGridView.Columns[8].HeaderCell.Value = "상품주소";

                int count = 1;

                foreach (var r in collection)
                {

                    dataGridView.Rows.Add(r.pageNo + "페이지 " + r.rank + "위",
                    r.oldPageNo + "페이지 " + r.oldrank + "위",
                    r.productName,
                    r.productNo,
                    r.categoryName,
                    r.similarity + "% ",
                    r.relevance + "% ",
                    r.hitRank,
                    r.productUrl);

                }
            }
        }


        public void SetAllDataRankGrid(object msgData, DataGridView dataGridView)
        {
            if (dataGridView.InvokeRequired)
            {
                DataGrid9 call = new DataGrid9(SetAllDataRankGrid);
                this.Invoke(call, msgData, dataGridView);
            }
            else
            {
                dataGridView.Rows.Clear();

                List<RankingList> collection = msgData as List<RankingList>;

                dataGridView.Rows.Clear();

                dataGridView.ColumnHeadersVisible = true;
                dataGridView.RowHeadersVisible = false;
                DataGridViewCellStyle columnHeaderStyle = new DataGridViewCellStyle();
                columnHeaderStyle.Font = new Font("Veradna", 10, FontStyle.Bold);
                columnHeaderStyle.BackColor = Color.Beige;
                dataGridView.ColumnHeadersDefaultCellStyle = columnHeaderStyle;
                dataGridView.ColumnCount = 13;
                dataGridView.Columns[0].HeaderCell.Value = "페이지별순위";
                dataGridView.Columns[1].HeaderCell.Value = "분류";
                dataGridView.Columns[2].HeaderCell.Value = "몰이름";
                dataGridView.Columns[3].HeaderCell.Value = "상품명";
                dataGridView.Columns[4].HeaderCell.Value = "상품번호";
                dataGridView.Columns[5].HeaderCell.Value = "카테고리";
                dataGridView.Columns[6].HeaderCell.Value = "리뷰수";
                dataGridView.Columns[7].HeaderCell.Value = "구매건수";
                dataGridView.Columns[8].HeaderCell.Value = "7일내구매근사치";
                dataGridView.Columns[9].HeaderCell.Value = "키워드유사성";
                dataGridView.Columns[10].HeaderCell.Value = "키워드관련성";
                dataGridView.Columns[11].HeaderCell.Value = "랭킹";
                dataGridView.Columns[12].HeaderCell.Value = "상품주소";

                int count = 1;

                foreach (var r in collection)
                {

                    dataGridView.Rows.Add(r.pageNo + "페이지 " + r.rank + "위",
                    r.adYn,
                    r.mallName,
                    r.productName,
                    r.productNo,
                    r.categoryName,
                    r.reviewCountSum,
                    r.purchaseCnt,
                    r.daysSaleSum7,
                    r.similarity + "% ",
                    r.relevance + "% ",
                    r.hitRank,
                    r.productUrl
                    );

                }
            }
        }


        public void SetDataGridClear2()
        {
            dataGridView2.Rows.Clear();
            dataGridView3.Rows.Clear();
            dataGridView8.Rows.Clear();
        }


        public void SetDataGridClear()
        {
            dataGridView1.Rows.Clear();
            dataGridView4.Rows.Clear();
            dataGridView6.Rows.Clear();
            dataGridView5.Rows.Clear();
            dataGridView7.Rows.Clear();
        }


        private void dataGridView6_SortCompare(object sender, DataGridViewSortCompareEventArgs e)
        {
            if (e.Column.Index == 1 || e.Column.Index == 2 || e.Column.Index == 9|| e.Column.Index == 10)    // 정렬할 컬럼의 이름

            {

                int a = int.Parse(e.CellValue1.ToString().Replace(",", "")), b = int.Parse(e.CellValue2.ToString().Replace(",", ""));

                e.SortResult = a.CompareTo(b);

                e.Handled = true;

            }
        }

        private void dataGridView7_SortCompare(object sender, DataGridViewSortCompareEventArgs e)
        {
            if (e.Column.Index == 2 || e.Column.Index == 3)    // 정렬할 컬럼의 이름

            {

                int a = int.Parse(e.CellValue1.ToString().Replace(",", "")), b = int.Parse(e.CellValue2.ToString().Replace(",", ""));

                e.SortResult = a.CompareTo(b);

                e.Handled = true;

            }
        }

        private void dataGridView4_SortCompare(object sender, DataGridViewSortCompareEventArgs e)
        {
            if (e.Column.Index == 1)    // 정렬할 컬럼의 이름

            {

                int a = int.Parse(e.CellValue1.ToString().Replace(",", "")), b = int.Parse(e.CellValue2.ToString().Replace(",", ""));

                e.SortResult = a.CompareTo(b);

                e.Handled = true;

            }
        }

        private void g1_btn_Click(object sender, EventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Title = "Save as Excel File";
            sfd.Filter = "Excel Files(2003)|*.xls|Excel Files(2007)|*.xlsx";
            sfd.FileName = "연관 검색어";
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                dataGridView_ExportToExcel(sfd.FileName, dataGridView6);
            }
        }

        private void g2_btn_Click(object sender, EventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Title = "Save as Excel File";
            sfd.Filter = "Excel Files(2003)|*.xls|Excel Files(2007)|*.xlsx";
            sfd.FileName = "쇼핑 연관";
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                dataGridView_ExportToExcel(sfd.FileName, dataGridView5);
            }
        }

        private void g3_btn_Click(object sender, EventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Title = "Save as Excel File";
            sfd.Filter = "Excel Files(2003)|*.xls|Excel Files(2007)|*.xlsx";
            sfd.FileName = "상품명";
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                dataGridView_ExportToExcel(sfd.FileName, dataGridView4);
            }
        }

        private void g4_btn_Click(object sender, EventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Title = "Save as Excel File";
            sfd.Filter = "Excel Files(2003)|*.xls|Excel Files(2007)|*.xlsx";
            sfd.FileName = "SEO";
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                dataGridView_ExportToExcel(sfd.FileName, dataGridView1);
            }
        }

        private void g5_btn_Click(object sender, EventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Title = "Save as Excel File";
            sfd.Filter = "Excel Files(2003)|*.xls|Excel Files(2007)|*.xlsx";
            sfd.FileName = "쇼핑몰정보";
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                dataGridView_ExportToExcel(sfd.FileName, dataGridView7);
            }

        }


        private void dataGridView_ExportToExcel(string fileName, DataGridView dgv)
        {
            Excel.Application excelApp = new Excel.Application();
            if (excelApp == null)
            {
                MessageBox.Show("엑셀이 설치되지 않았습니다");
                return;
            }
            Excel.Workbook wb = excelApp.Workbooks.Add(true);
            Excel._Worksheet workSheet = wb.Worksheets.get_Item(1) as Excel._Worksheet;
            workSheet.Name = "정보";

            if (dgv.Rows.Count == 0)
            {
                MessageBox.Show("출력할 데이터가 없습니다");
                return;
            }

            // storing header part in Excel  
            for (int i = 1; i < dgv.Columns.Count + 1; i++)
            {
                workSheet.Cells[1, i] = dgv.Columns[i - 1].HeaderText;
            }
            // storing Each row and column value to excel sheet  
            for (int i = 0; i < dgv.Rows.Count - 1; i++)
            {
                for (int j = 0; j < dgv.Columns.Count; j++)
                {
                    workSheet.Cells[i + 2, j + 1] = dgv.Rows[i].Cells[j].Value.ToString();
                }
            }

            // 엑셀 2003 으로만 저장이 됨
            wb.SaveAs(fileName, Excel.XlFileFormat.xlWorkbookNormal, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            wb.Close(Type.Missing, Type.Missing, Type.Missing);
            excelApp.Quit();
            releaseObject(excelApp);
            releaseObject(workSheet);
            releaseObject(wb);

            SetListBox("엑셀파일 생성을 완료하였습니다.");
        }


        private static void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception e)
            {
                obj = null;
            }
            finally
            {
                GC.Collect();
            }
        }

        private void keyword()
        {

            string keyword = Uri.EscapeUriString(textBox1.Text);



            string url = "https://search.shopping.naver.com/search/all.nhn?query=" + keyword + "&cat_id=&frm=NVSHATC";
            string murl = "https://msearch.shopping.naver.com/search/all?query=" + keyword + "&cat_id=&frm=NVSHATC";

            var additionalHeaders = "User-Agent:Mozilla/5.0 (Windows Phone 10.0; Android 6.0.1; " +
                                        "Microsoft; Lumia 950 XL Dual SIM) AppleWebKit/537.36 (KHTML, like Gecko) " +
                                        "Chrome/52.0.2743.116 Mobile Safari/537.36 Edge/15.15063\r\n";

            this.webBrowser3.Navigate(murl, null, null, additionalHeaders);
            webBrowser4.Navigate(url);
        }

        // 쇼핑검색
        private void button1_Click(object sender, EventArgs e)
        {

            keyword();

        }

        private void textBox1_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                keyword();
            }


        }

        private bool timerYn = false;

        private void button2_Click(object sender, EventArgs e)
        {
            listBox2.Items.Clear();
            SetDataGridClear2();


            if (textBox2.Text.Length < 1)
            {
                MessageBox.Show("쇼핑몰명을 넣어주세요!", "경고", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (textBox3.Text.Length < 1)
            {
                MessageBox.Show("키워드를 넣어주세요!", "경고", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (textBox4.Text.Length < 1)
            {
                MessageBox.Show("광고비를 넣어주세요! 없으면 0 을 넣어주세요.", "경고", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            int intervalTime = 0;
            // 타이머 반복작업

            if (comboBox1.SelectedIndex == 0)
            {
                intervalTime = 0;
            }
            else if (comboBox1.SelectedIndex == 1)
            {
                intervalTime = 30 * 60 * 1000; // 30분
            }
            else if (comboBox1.SelectedIndex == 2)
            {
                intervalTime = 60 * 60 * 1000; // 1시간
            }
            else if (comboBox1.SelectedIndex == 3)
            {
                intervalTime = 90 * 60 * 1000; // 1시간 30분
            }
            else if (comboBox1.SelectedIndex == 4)
            {
                intervalTime = 120 * 60 * 1000; // 2시간
            }
            else if (comboBox1.SelectedIndex == 5)
            {
                intervalTime = 150 * 60 * 1000; // 2시간 30분
            }
            else if (comboBox1.SelectedIndex == 6)
            {
                intervalTime = 180 * 60 * 1000; // 3시간
            }
            else if (comboBox1.SelectedIndex == 7)
            {
                intervalTime = 210 * 60 * 1000; // 3시간
            }
            else
            {
                intervalTime = 0;
            }

            if (intervalTime > 0)
            {
                if (timerYn == false)
                {
                    timerYn = true;
                    timer.Interval = intervalTime;
                    timer.Start();
                    button2.Text = "중지";
                    Thread t1 = new Thread(new ThreadStart(AdDataReturn));
                    t1.Start();

                }
                else
                {
                    SetListBox2("반복작업을 중지합니다.");
                    timerYn = false;
                    timer.Stop();
                    button2.Text = "시작";
                }
            }
            else
            {

                Thread t1 = new Thread(new ThreadStart(AdDataReturn));
                t1.Start();
            }
        }

        private void Timer_Tick(object sender, EventArgs e)
        {
            Thread t1 = new Thread(new ThreadStart(AdDataReturn));
            t1.Start();

            SetListBox2("데이터 조회시간:"+DateTime.Now.ToString());
            
        }

        public void AdDataReturn()
        {
            SetButton2(false);
            GridResultData2 gridResultData2 = new GridResultData2();
            SetListBox2("데이터를 조회하기 위해 초기화 중입니다.");
            bizranking outData = new bizranking();
            gridResultData2 = outData.SamartStoreRankingSearch(textBox3.Text, textBox2.Text, textBox4.Text);

            SetDataAdRankGrid(gridResultData2.AdRankingRefGrid, dataGridView2); //전체 연관 검색어 리스트
            SetDataRankGrid(gridResultData2.NonAdRankingRefGrid, dataGridView3);
            SetAllDataRankGrid(gridResultData2.AllRankingRefGrid, dataGridView8);

            SetButton2(true);
        }

        private void g6_btn_Click(object sender, EventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Title = "Save as Excel File";
            sfd.Filter = "Excel Files(2003)|*.xls|Excel Files(2007)|*.xlsx";
            sfd.FileName = "광고상품순위";
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                dataGridView_ExportToExcel(sfd.FileName, dataGridView2);
            }

        }

        private void g7_btn_Click(object sender, EventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Title = "Save as Excel File";
            sfd.Filter = "Excel Files(2003)|*.xls|Excel Files(2007)|*.xlsx";
            sfd.FileName = "일반상품순위";
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                dataGridView_ExportToExcel(sfd.FileName, dataGridView3);
            }

        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start("https://www.photopea.com/");
        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start("https://www.remove.bg/ko");
        }

        private void linkLabel3_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start("https://www.canva.com/");
        }

        private void linkLabel4_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start("https://www.mangoboard.net/ ");
        }

        private void linkLabel5_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start("http://x.photoscape.org/");
        }

        private void linkLabel6_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start("https://blog.naver.com/darkwalk77");
        }

        private void linkLabel7_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start("https://modu-print.tistory.com/");
        }

        private void linkLabel8_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start("https://gongu.copyright.or.kr/freeFontEvent.html");
        }

        private void linkLabel9_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start("https://open.kakao.com/o/g4wjxW1b");
        }

        private void dataGridView7_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridView7.CurrentCell == null || dataGridView7.CurrentCell.Value == null || e.RowIndex == -1)
            {
                return;
            }

            if (dataGridView7.CurrentCell.ColumnIndex.Equals(5))
            {
                string linkurl = dataGridView7.CurrentCell.Value.ToString();
                System.Diagnostics.Process.Start(linkurl);

            }
        }

        private void linkLabel10_LinkClicked_1(object sender, LinkLabelLinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start("http://www.parcelman.kr/");
        }

        /// <summary>
        /// 인스타그램 태그 검색
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void instarTagBtn_Click(object sender, EventArgs e)
        {

            if (instatagBox.Text.Length < 1)
            {
                MessageBox.Show("키워드를 넣어주세요!", "경고", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            // 그리드 클리어
            instaDataGridView.Rows.Clear();

            Thread t1 = new Thread(new ThreadStart(instargramTagDataSet));
            t1.Start();
        }

        public void instargramTagDataSet()
        {
            SetInstaButton(false);
            List<InstagramTagWordList> InstaWordList = new List<InstagramTagWordList>();
            InstagramAPI instagram = new InstagramAPI();
            string result = string.Concat(instatagBox.Text.Where(c => !char.IsWhiteSpace(c)));
            string input = Regex.Replace(result, @"[^a-zA-Z0-9가-힣_]", "", RegexOptions.Singleline);
            InstaWordList = instagram.InstagramJsonDataSet(input);

            //해시태그 
            SetInstaDataGrid(InstaWordList, instaDataGridView);
            //webBrowser6.Navigate("https://www.instagram.com/explore/tags/" + input);

            SetInstaButton(true);
        }

        private void instaDataGridView_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (instaDataGridView.CurrentCell == null || instaDataGridView.CurrentCell.Value == null || e.RowIndex == -1)
            {
                return;
            }

            if (instaDataGridView.CurrentCell.ColumnIndex.Equals(0))
            {
                string instatag = instaDataGridView.CurrentCell.Value.ToString();
                string result = string.Concat(instatag.Where(c => !char.IsWhiteSpace(c)));
                string input = Regex.Replace(result, @"[^a-zA-Z0-9가-힣_]", "", RegexOptions.Singleline);
                //webBrowser6.Navigate("https://www.instagram.com/explore/tags/" + input);

            }
        }

        private void instaDataGridView_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (instaDataGridView.CurrentCell == null || instaDataGridView.CurrentCell.Value == null || e.RowIndex == -1)
            {
                return;
            }

            if (instaDataGridView.CurrentCell.ColumnIndex.Equals(0))
            {
                string instatag = instaDataGridView.CurrentCell.Value.ToString();
                string result = string.Concat(instatag.Where(c => !char.IsWhiteSpace(c)));
                string input = Regex.Replace(result, @"[^a-zA-Z0-9가-힣_]", "", RegexOptions.Singleline);
                instatagBox.Text = input;

                if (instatagBox.Text.Length < 1)
                {
                    MessageBox.Show("키워드를 넣어주세요!", "경고", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                // 그리드 클리어
                instaDataGridView.Rows.Clear();

                Thread t1 = new Thread(new ThreadStart(instargramTagDataSet));
                t1.Start();

            }

        }

        private void catebnt_Click(object sender, EventArgs e)
        {
            Thread t1 = new Thread(new ThreadStart(categoryDataSet));
            t1.Start();
        }

        public void categoryDataSet()
        {
            List<CategoryList> categoryLists = new List<CategoryList>();
            NaverShoppingCrawler naverShoppingCrawler = new NaverShoppingCrawler();
            //categoryLists = naverShoppingCrawler.CategoryRsultText(CateSBox.Text);
        }

        private void button3_Click(object sender, EventArgs e)
        {

            if (DataResult.AdRefGrid.Count != 0)
            {
                int MonthlyPcQcCntdata = 0;
                int MonthlyMobileQcCntdata = 0;
                int SellPrdQcCntdata = 0;

                if (string.IsNullOrEmpty(textBox5.Text))
                {
                    MonthlyPcQcCntdata = 0;
                }
                else
                {
                    MonthlyPcQcCntdata = Convert.ToInt32(textBox5.Text);
                }

                if (string.IsNullOrEmpty(textBox6.Text))
                {
                    MonthlyMobileQcCntdata = 0;
                }
                else
                {
                    MonthlyMobileQcCntdata = Convert.ToInt32(textBox6.Text);
                }

                if (string.IsNullOrEmpty(textBox7.Text))
                {
                    SellPrdQcCntdata = 1000000000;
                }
                else
                {
                    SellPrdQcCntdata = Convert.ToInt32(textBox7.Text);
                }

                List<KeyWordResult> items = new List<KeyWordResult>();

                var shopResult = from a in DataResult.AdRefGrid where

                                // 조건문
                                Convert.ToInt32(a.MonthlyPcQcCnt) >= MonthlyPcQcCntdata

                                && Convert.ToInt32(a.MonthlyMobileQcCnt) >= MonthlyMobileQcCntdata

                                && Convert.ToInt32(a.SellPrdQcCnt) <= SellPrdQcCntdata



                                 select new KeyWordResult
                                 {

                                     RelKeyword = a.RelKeyword,
                                     SellPrdQcCnt = a.SellPrdQcCnt,
                                     MonthlyPcQcCnt = a.MonthlyPcQcCnt,
                                     MonthlyMobileQcCnt = a.MonthlyMobileQcCnt,
                                     MonthlyAvePcClkCnt = a.MonthlyAvePcClkCnt,
                                     MonthlyAveMobileClkCnt = a.MonthlyAveMobileClkCnt,
                                     MonthlyAvePcCtr = a.MonthlyAvePcCtr,
                                     MonthlyAveMobileCtr = a.MonthlyAveMobileCtr,
                                     PlAvgDepth = a.PlAvgDepth,
                                     CompIdx = a.CompIdx,
                                     ShopResult = a.ShopResult
                                 };

                items.AddRange(shopResult);


                SetDataGrid(items, dataGridView6); //전체 연관 검색어 리스트  
            }
            else
            {
                MessageBox.Show("키워드 검색후 조건검색을 하십시오.", "경고", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }

        private void textBox5_KeyPress(object sender, KeyPressEventArgs e)
        {
            //숫자만 입력되도록 필터링
            if (!(char.IsDigit(e.KeyChar) || e.KeyChar == Convert.ToChar(Keys.Back)))    //숫자와 백스페이스를 제외한 나머지를 바로 처리
            {
                e.Handled = true;
            }
        }

        private void textBox6_KeyPress(object sender, KeyPressEventArgs e)
        {
            //숫자만 입력되도록 필터링
            if (!(char.IsDigit(e.KeyChar) || e.KeyChar == Convert.ToChar(Keys.Back)))    //숫자와 백스페이스를 제외한 나머지를 바로 처리
            {
                e.Handled = true;
            }
        }

        private void textBox7_KeyPress(object sender, KeyPressEventArgs e)
        {
            //숫자만 입력되도록 필터링
            if (!(char.IsDigit(e.KeyChar) || e.KeyChar == Convert.ToChar(Keys.Back)))    //숫자와 백스페이스를 제외한 나머지를 바로 처리
            {
                e.Handled = true;
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Title = "Save as Excel File";
            sfd.Filter = "Excel Files(2003)|*.xls|Excel Files(2007)|*.xlsx";
            sfd.FileName = "전체상품순위";
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                dataGridView_ExportToExcel(sfd.FileName, dataGridView8);
            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            
        }
    }
}
