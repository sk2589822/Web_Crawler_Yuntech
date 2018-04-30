using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using HtmlAgilityPack;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.Threading;
using System.Data.OleDb;
using MySql.Data.MySqlClient;

namespace Web_Crawler
{
    public partial class Form1 : Form
    {
        static string YunTech_url = "http://events.yuntech.edu.tw/";
        static string[] Web_field = { "主辦單位:", "主持人:", "時間:", "活動描述:", "地點:", "連繫:", "備註" };
        static string[] Excel_field = { "", "類型", "名稱", "主辦單位", "主持人", "時間", "活動描述", "地點", "連繫", "備註" };
        static string[] html_path = { "http://webapp.yuntech.edu.tw/WebMSS/bulletin.aspx?deptcode=AA",
                                      "http://asx.yuntech.edu.tw/index.php?option=com_content&task=News&id=7&limit=15&limitstart=0",
                                      "http://ags.yuntech.edu.tw/index.php?option=com_content&task=News&id=7&limit=15&limitstart=0",
                                      "http://asp.yuntech.edu.tw/index.php?option=com_content&task=News&id=7&limit=25&limitstart=0",
                                      "http://aex.yuntech.edu.tw/index.php?option=com_content&task=News&id=7&limit=15&limitstart=0",
                                      "http://tdx.yuntech.edu.tw/index.php?option=com_content&task=News&id=7&limit=15&limitstart=0",
                                      "http://lc.yuntech.edu.tw/index.php?option=com_content&task=News&id=7&limit=15&limitstart=0",
                                      "http://libweb.yuntech.edu.tw/news/index.php?pg=",
                                      "http://ttx.yuntech.edu.tw/最新消息/page/0",};
        static string[] Office_name = { "教務處", "學務處", "總務處", "體育處", "人事處", "國際事務處", "語言中心", "圖書館", "研發處"};
        static public int[] Months = { 0, 31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31 };
        int[] anno_saved = { 0, 0, 0, 0, 0 };
        Excel.Application App;
        Excel.Workbook book;
        static WaitHandle[] waitHandles = null;
        static object _lock = new object();


        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        public class save_data
        {
            public string file_name;
            public int year;
            public int month;
            public int day;
            public Excel.Worksheet sheet;
            public AutoResetEvent _WaitHandle;
            public AutoResetEvent WaitHandle
            {
                get { return _WaitHandle; }
                set { _WaitHandle = value; }
            }
            public save_data(string name, int y, int m, int d, Excel.Worksheet sht)
            {
                file_name = name;
                year = y;
                month = m;
                day = d;
                sheet = sht;
                WaitHandle = new AutoResetEvent(false);
            }
        }

        private void Get_url(HtmlNode node, Excel.Worksheet sheet, int row)
        {
            string url = node.InnerHtml;
            int start = url.IndexOf("href=");
            int end = url.IndexOf(";\"", start);
            url = url.Substring(start + 6, end - start - 5);
            url = url.Replace("amp;", "");

            Get_Detail(YunTech_url + url, sheet, row);
        }

        private void Get_Detail(string event_url, Excel.Worksheet sheet, int row)
        {

            HtmlWeb webClient = new HtmlWeb();
            HtmlAgilityPack.HtmlDocument doc = webClient.Load(event_url);
            HtmlAgilityPack.HtmlDocument docStockContext = new HtmlAgilityPack.HtmlDocument();

            docStockContext.LoadHtml(doc.DocumentNode.SelectSingleNode(
            "/html[1]/body[1]/div[1]/div[5]/div[1]/div[3]/div[1]/div[1]/div[1]").InnerHtml);

            HtmlNodeCollection nodes = docStockContext.DocumentNode.SelectNodes("table[1]/tbody[1]/tr/td");

            for (int i = 0; i < nodes.Count; i += 2)
            {
                int col = 0;
                string str = nodes[i].InnerText;
                str = str.Replace("\t", "");
                str = str.Replace("\n", "");
                str = str.Replace("\r", "");
                str = str.Replace("　", "");
                str = str.Replace(" ", "");
                while (str != Web_field[col++] && col < 6) ;
                if (col < 6)
                {
                    str = nodes[i + 1].InnerText;
                    str = str.Replace("\t", "");
                    str = str.Replace("\n", "");
                    str = str.Replace("\r", "");
                    str = str.Replace("　", "");
                    str = str.Replace(" ", "");
                    sheet.Cells[row, col + 2] = str;
                }
            }
        }

        public delegate void myUICallBack(string myStr, TextBox ctl);

        public void myUI(string myStr, TextBox ctl)
        {
            if (this.InvokeRequired)
            {
                myUICallBack myUpdate = new myUICallBack(myUI);
                this.Invoke(myUpdate, myStr, ctl);
            }
            else
            {
                ctl.Text += myStr;
                ctl.SelectionStart = ctl.Text.Length;
                ctl.ScrollToCaret();
                ctl.Focus();
            }
        }

        private void Get_day_data_Click(object sender, EventArgs e)
        {
            int year, month, day;
            year = dateTimePicker_day.Value.Year;
            month = dateTimePicker_day.Value.Month;
            day = dateTimePicker_day.Value.Day;
            string Event_url = "?&y=" + year + "&m=" + month + "&d=" + day + "&";

            HtmlWeb webClient = new HtmlWeb();
            HtmlAgilityPack.HtmlDocument doc = webClient.Load(YunTech_url + Event_url);
            HtmlAgilityPack.HtmlDocument docStockContext = new HtmlAgilityPack.HtmlDocument();
            docStockContext.LoadHtml(doc.DocumentNode.SelectSingleNode("/html[1]/body[1]/div[1]/div[5]/div[1]/div[3]/div[1]").InnerHtml);
            HtmlNodeCollection event_check = docStockContext.DocumentNode.SelectNodes("table");

            if (event_check == null)
            {
                textBox1.Text = "今天沒有新的或是正在進行中的活動";
            }
            else if (event_check.Count == 1)
            {
                HtmlNode title = docStockContext.DocumentNode.SelectSingleNode("table[1]/thead/tr/th");
                HtmlNodeCollection nodes = docStockContext.DocumentNode.SelectNodes("table[1]/tbody/tr/td");

                if (title.InnerText == "時間")
                {
                    textBox1.Text = "今天的新活動：\r\n";
                }
                else
                {
                    textBox1.Text = "今天的進行中活動：\r\n";
                }

                bool endl = false;
                foreach (HtmlNode node in nodes)
                {
                    string[] values = node.InnerText.Split(' ');
                    if (endl)
                    {
                        textBox1.Text += values[0] + "    " + values[1] + "\r\n";
                    }
                    else
                    {
                        textBox1.Text += node.InnerText + "    ";
                    }
                    endl = !endl;
                }
            }
            else
            {
                HtmlNodeCollection new_nodes = docStockContext.DocumentNode.SelectNodes("table[1]/tbody/tr/td");
                HtmlNodeCollection ongoing_nodes = docStockContext.DocumentNode.SelectNodes("table[2]/tbody/tr/td");

                textBox1.Text = "今天的新活動：\r\n";
                bool endl = false;
                foreach (HtmlNode node in new_nodes)
                {
                    string[] values = node.InnerText.Split(' ');

                    if (endl)
                    {
                        textBox1.Text += values[0] + "    " + values[1] + "\r\n";
                    }
                    else
                    {
                        textBox1.Text += node.InnerText + "    ";
                    }
                    endl = !endl;
                }

                textBox1.Text += "\r\n今天的進行中活動：\r\n";
                endl = false;
                foreach (HtmlNode node in ongoing_nodes)
                {
                    string[] values = node.InnerText.Split(' ');
                    if (endl)
                    {
                        textBox1.Text += values[0] + "    " + values[1] + "\r\n";
                    }
                    else
                    {
                        textBox1.Text += node.InnerText + "    ";
                    }
                    endl = !endl;
                }
            }
        }

        public void Get_Calendar_day(object data)
        {
            AutoResetEvent reset = (data as save_data).WaitHandle;
            string file_name = (data as save_data).file_name;
            int year = (data as save_data).year;
            int month = (data as save_data).month;
            int day = (data as save_data).day;
            string day_str = day.ToString().PadLeft(2, '0');
            string Excel_path = @"E:\Web Crawler\" + file_name + ".xlsx";
            Excel.Worksheet sheet = (data as save_data).sheet;

            for (int i = 1; i <= 9; i++)
            {
                sheet.Cells[1, i] = Excel_field[i];
            }

            string Event_url = "?&y=" + year + "&m=" + month + "&d=" + day + "&";
            HtmlWeb webClient = new HtmlWeb();
            HtmlAgilityPack.HtmlDocument doc = webClient.Load(YunTech_url + Event_url);
            HtmlAgilityPack.HtmlDocument docStockContext = new HtmlAgilityPack.HtmlDocument();
            docStockContext.LoadHtml(doc.DocumentNode.SelectSingleNode("/html[1]/body[1]/div[1]/div[5]/div[1]/div[3]/div[1]").InnerHtml);
            HtmlNodeCollection event_check = docStockContext.DocumentNode.SelectNodes("table");

            int row = 2;

            if (event_check != null)
            {
                if (event_check.Count == 1)
                {
                    HtmlNode title = docStockContext.DocumentNode.SelectSingleNode("table[1]/thead/tr/th");
                    HtmlNodeCollection nodes = docStockContext.DocumentNode.SelectNodes("table[1]/tbody/tr/td");
                    string str_category = "";

                    if (title.InnerText == "時間")
                    {
                        str_category = "新活動";
                    }
                    else
                    {
                        str_category = "進行中活動";
                    }

                    for (int i = 0; i < nodes.Count; i++)
                    {
                        string[] values = nodes[++i].InnerText.Split(' ');
                        sheet.Cells[row, 1] = str_category;
                        sheet.Cells[row, 2] = values[0];
                        Get_url(nodes[i], sheet, row++);
                        /*  if (nodes.Count - i == 1)
                            {
                                myUI(file_name + "_" + day_str + "已抓取 (" + (i / 2 + 1) + "/" + (nodes.Count / 2) + ") 筆...抓取完畢\r\n", textBox1);
                            }
                            else
                            {
                                myUI(file_name + "_" + day_str + "已抓取 (" + (i / 2 + 1) + "/" + (nodes.Count / 2) + ") 筆\r\n", textBox1);
                            }
                         */
                    }
                }
                else
                {
                    HtmlNodeCollection new_nodes = docStockContext.DocumentNode.SelectNodes("table[1]/tbody/tr/td");
                    HtmlNodeCollection ongoing_nodes = docStockContext.DocumentNode.SelectNodes("table[2]/tbody/tr/td");

                    int count = (new_nodes.Count + ongoing_nodes.Count) / 2;
                    for (int i = 0; i < new_nodes.Count; i++)
                    {
                        string[] values = new_nodes[++i].InnerText.Split(' ');
                        sheet.Cells[row, 1] = "新活動";
                        sheet.Cells[row, 2] = values[0];
                        Get_url(new_nodes[i], sheet, row++);
                        //myUI(file_name + "_" + day_str + "已抓取 (" + (i / 2 + 1) + "/" + count + ") 筆\r\n", textBox1);
                    }
                    for (int i = 0; i < ongoing_nodes.Count; i++)
                    {
                        string[] values = ongoing_nodes[++i].InnerText.Split(' ');
                        sheet.Cells[row, 1] = "進行中活動";
                        sheet.Cells[row, 2] = values[0];
                        Get_url(ongoing_nodes[i], sheet, row++);
                        /*  if (ongoing_nodes.Count - i == 1)
                            {
                                myUI(file_name + "_" + day_str + "已抓取 (" + ((i + new_nodes.Count) / 2 + 1) + "/" + count + ") 筆...抓取完畢\r\n", textBox1);
                            }
                            else
                            {
                                myUI(file_name + "_" + day_str + "已抓取 (" + ((i + new_nodes.Count) / 2 + 1) + "/" + count + ") 筆\r\n", textBox1);
                            }
                          */
                    }
                }
            }
            reset.Set();
        }

        private void Get_Calendar_func(int year, int month)
        {
            string file_name = year.ToString() + "_" + month.ToString().PadLeft(2, '0');
            string Excel_path = @"E:\Web Crawler\" + file_name + ".xlsx";
            App = new Excel.Application();
            book = App.Workbooks.Add(true);

            int max = Months[month];
            if (month == 2 && DateTime.IsLeapYear(year))
            {
                max++;
            }

            for (int i = 1; i <= max; i++)
            {
                if (i > 1)
                {
                    book.Sheets.Add(Missing.Value, book.Sheets[i - 1], Missing.Value, Missing.Value);
                }
                Excel.Worksheet sheet = (Excel.Worksheet)book.Sheets[i];
                sheet.Name = (i.ToString().PadLeft(2, '0'));
                sheet.Application.DisplayAlerts = false;
                sheet.Application.AlertBeforeOverwriting = false;
            }
            Thread[] thread = new Thread[35];
            ThreadPool.SetMaxThreads(10, 10);
            waitHandles = new WaitHandle[max + 1];
            System.Threading.WaitCallback waitCallback = new WaitCallback(Get_Calendar_day);
            waitHandles[0] = new AutoResetEvent(true);
            for (int day = 1; day <= max; day++)
            {
                Excel.Worksheet sheet = (Excel.Worksheet)book.Sheets[day];
                save_data data = new save_data(file_name, year, month, day, sheet);
                waitHandles[day] = data.WaitHandle;
                ThreadPool.QueueUserWorkItem(new WaitCallback(waitCallback), data);
            }
            for (int i = max + 1; i < waitHandles.Count(); i++)
            {
                waitHandles[i] = new AutoResetEvent(true);
            }

            WaitHandle.WaitAll(waitHandles);
            book.SaveCopyAs(Excel_path);
            book.Close(0);
            App.Quit();
        }

        private void Get_Calendar_Click(object sender, EventArgs e)
        {
            int year_start, month_start, year_end, month_end;
            year_start = dateTimePicker_month_1.Value.Year;
            month_start = dateTimePicker_month_1.Value.Month;
            year_end = dateTimePicker_month_2.Value.Year;
            month_end = dateTimePicker_month_2.Value.Month;
            textBox1.Text += "取得行事曆中...\r\n";
            for (int y = year_start, m = month_start; y * 100 + m <= year_end * 100 + month_end; m++)
            {
                if (m > 12)
                {
                    m = 1;
                    y++;
                }
                Get_Calendar_func(y, m);
            }
            textBox1.Text += "成功取得行事曆資料\r\n";
        }

        public void Upload_Calendar_func(string name)
        {
            string Excel_path = @"E:\Web Crawler\" + name + ".xlsx";
            App = new Excel.Application();
            book = App.Workbooks.Open(Excel_path, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                                  Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
            int year = Int32.Parse(name.Substring(0, 4));
            int month = Int32.Parse(name.Substring(name.Length - 2, 2));
            int max = Months[month];
            if (month == 2 && DateTime.IsLeapYear(year))
            {
                max++;
            }
            for (int p = 1; p <= max; p++)
            {
                Excel.Worksheet sheet = (Excel.Worksheet)book.Worksheets.get_Item(p);

                int rows_count = sheet.UsedRange.Cells.Rows.Count;

                Excel.Range rng = sheet.Cells.get_Range("A2", "K" + rows_count);

                for (int i = 2; i <= rows_count; i++)
                {
                    string dbHost = "localhost";//資料庫位址
                    string dbUser = "root";//資料庫使用者帳號
                    string dbPass = "";//資料庫使用者密碼
                    string dbName = "csv_db";//資料庫名稱

                    string connStr = "server=" + dbHost + ";uid=" + dbUser + ";pwd=" + dbPass + ";database=" + dbName + ";CharSet =" + "utf8";
                    MySqlConnection conn = new MySqlConnection(connStr);
                    MySqlCommand command = conn.CreateCommand();
                    conn.Open();

                    string title;
                    try
                    {
                        title = sheet.Cells[i, 2].Value2.Replace("'", "’");
                    }
                    catch
                    {
                        continue;
                    }
                    command.CommandText = "Insert into " + name + "(date,type,name,host,hoster,time,descr,location,connect,remark) values('" + p.ToString() + "','" + sheet.Cells[i, 1].Value2 + "','" + sheet.Cells[i, 2].Value2.Replace("'","’") + "','" + sheet.Cells[i, 3].Value2 + "','" + sheet.Cells[i, 4].Value2 + "','" + sheet.Cells[i, 5].Value2 + "','" + sheet.Cells[i, 6].Value2 + "','" + sheet.Cells[i, 7].Value2 + "','" + sheet.Cells[i, 8].Value2 + "','" + sheet.Cells[i, 9].Value2 + "')";
                    command.ExecuteNonQuery();

                    conn.Close();

                }
            }
            book.Close(0);
            App.Quit();
        }

        private void Upload_Calendar_Click(object sender, EventArgs e)
        {
            string name = Upload_Text.Text;
            if (name.Length == 7)
            {
                textBox1.Text += "上傳" + name + "月份行事曆中......\r\n";
                Upload_Calendar_func(name);
            }
            else
            {
                int year_strat = Convert.ToInt32(name.Substring(0, 4));
                int month_strat = Convert.ToInt32(name.Substring(5, 2));
                int year_end = Convert.ToInt32(name.Substring(11, 4));
                int month_end = Convert.ToInt32(name.Substring(16, 2));
                textBox1.Text += "上傳" + year_strat + "_" + month_strat.ToString().PadLeft(2, '0') + "月份至" + year_end + "_" + month_end.ToString().PadLeft(2, '0') + "月份行事曆中......\r\n";
                for (int y = year_strat, m = month_strat; y * 100 + m <= year_end * 100 + month_end; m++)
                {
                    if (m > 12)
                    {
                        m = 1;
                        y++;
                    }
                    name = y.ToString() + "_" + m.ToString().PadLeft(2, '0');
                    
                    Upload_Calendar_func(name);
                }
            }
            textBox1.Text += "上傳成功\r\n";
        }

        public void Create_Table_func(string name)
        {
            string dbHost = "localhost";//資料庫位址
            string dbUser = "root";//資料庫使用者帳號
            string dbPass = "";//資料庫使用者密碼
            string dbName = "csv_db";//資料庫名稱

            string connStr = "server=" + dbHost + ";uid=" + dbUser + ";pwd=" + dbPass + ";database=" + dbName;
            MySqlConnection conn = new MySqlConnection(connStr);
            MySqlCommand command = conn.CreateCommand();
            conn.Open();
            command.CommandText = "CREATE TABLE " + name + "(date VARCHAR(100) , type VARCHAR(100),name VARCHAR(100),host VARCHAR(100),hoster VARCHAR(100),time VARCHAR(100),descr VARCHAR(100),location VARCHAR(100),connect VARCHAR(100),remark VARCHAR(100))";
            command.ExecuteNonQuery();
            conn.Close();
        }

        private void Create_Table_Click(object sender, EventArgs e)
        {
            string name = Create_Text.Text;
            if (name.Length == 7)
            {
                Create_Table_func(name);
                textBox1.Text += "已建立" + name + "月份表格\r\n";
            }
            else
            {
                int year_strat = Convert.ToInt32(name.Substring(0, 4));
                int month_strat = Convert.ToInt32(name.Substring(5, 2));
                int year_end = Convert.ToInt32(name.Substring(11, 4));
                int month_end = Convert.ToInt32(name.Substring(16, 2));
                for (int y = year_strat, m = month_strat; y * 100 + m <= year_end * 100 + month_end; m++)
                {
                    if (m > 12)
                    {
                        m = 1;
                        y++;
                    }
                    name = y.ToString() + "_" + m.ToString().PadLeft(2, '0');
                    Create_Table_func(name);
                }
                textBox1.Text += "已上傳" + year_strat + "_" + month_strat.ToString().PadLeft(2, '0') + "月份至" + year_end + "_" + month_end.ToString().PadLeft(2, '0') + "月份表格\r\n";
            }
        }

        private void Delete_Table_Click(object sender, EventArgs e)
        {
            if (DialogResult.Yes == MessageBox.Show("您確定要刪除？", "Table刪除", MessageBoxButtons.YesNo, MessageBoxIcon.Question))
            {
                string dbHost = "localhost";//資料庫位址
                string dbUser = "root";//資料庫使用者帳號
                string dbPass = "";//資料庫使用者密碼
                string dbName = "csv_db";//資料庫名稱

                string connStr = "server=" + dbHost + ";uid=" + dbUser + ";pwd=" + dbPass + ";database=" + dbName;
                MySqlConnection conn = new MySqlConnection(connStr);
                MySqlCommand command = conn.CreateCommand();
                conn.Open();
                command.CommandText = "Drop table " + Delete_Text.Text + ";";
                command.ExecuteNonQuery();
                conn.Close();
                MessageBox.Show("已刪除！", "刪除table", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        public void Clear_Table_func(string name)
        {
            string dbHost = "localhost";//資料庫位址
            string dbUser = "root";//資料庫使用者帳號
            string dbPass = "";//資料庫使用者密碼
            string dbName = "csv_db";//資料庫名稱

            string connStr = "server=" + dbHost + ";uid=" + dbUser + ";pwd=" + dbPass + ";database=" + dbName;
            MySqlConnection conn = new MySqlConnection(connStr);
            MySqlCommand command = conn.CreateCommand();
            conn.Open();
            command.CommandText = "Truncate Table " + name + ";";
            // command.CommandText = "Drop table " + comboBox2.Text.ToString() + ";";
            command.ExecuteNonQuery();
            conn.Close();

        }

        private void Clear_Table_Click(object sender, EventArgs e)
        {
            if (DialogResult.Yes == MessageBox.Show("您確定要清除？", "Table清除", MessageBoxButtons.YesNo, MessageBoxIcon.Question))
            {
                Clear_Table_func(Clear_Text.Text);
                MessageBox.Show("已清除！", "清除table", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        public int Get_Sum(HtmlNode sum_tr, string start_word, string end_word)
        {
            string sum_str = sum_tr.InnerText.Trim();
            int start = sum_str.IndexOf(start_word);
            int end = sum_str.IndexOf(end_word, start);
            sum_str = sum_str.Substring(start + 1, end - start - 1).Trim();
            int sum = Int32.Parse(sum_str);
            return sum;
        }

        public void Get_Announcement_func()
        {
            string Excel_path = @"G:\我的文件\雲科\專題\Web Crawler_2016_08_28 + 4Website\Announcement.xlsx";
            App = new Excel.Application();
            book = App.Workbooks.Add(true);

            HtmlWeb webClient = new HtmlWeb();
            HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
            HtmlAgilityPack.HtmlDocument docStockContext = new HtmlAgilityPack.HtmlDocument();
            HtmlNodeCollection nodes;
            WebClient client = new WebClient();
            MemoryStream ms = new MemoryStream();
            String path, title, department, time, url;

            for (int i = 0; i < 9; i++)
            {
                if (i > 0)
                {
                    book.Sheets.Add(Missing.Value, book.Sheets[book.Sheets.Count], Missing.Value, Missing.Value);
                }
                Excel.Worksheet sheet = (Excel.Worksheet)book.Sheets[book.Sheets.Count];
                sheet.Name = Office_name[i];
                sheet.Columns[3].ColumnWidth = 12;
                sheet.Application.DisplayAlerts = false;
                sheet.Application.AlertBeforeOverwriting = false;
                sheet.Cells[1, 1] = "公告事項";
                sheet.Cells[1, 2] = "公告單位";
                sheet.Cells[1, 3] = "公告日期";
                sheet.Cells[1, 4] = "網址";


                if (i == 0)
                {
                    ms = new MemoryStream(client.DownloadData(html_path[i]));
                    doc.Load(ms, Encoding.UTF8);
                    docStockContext.LoadHtml(doc.DocumentNode.SelectSingleNode("/html[1]/body[1]/div[1]/table[1]/tr[1]/td[1]/div[1]/table[1]").InnerHtml);
                    nodes = docStockContext.DocumentNode.SelectNodes("/tr");

                    for (int j = 2; j <= nodes.Count(); j++)
                    {
                        path = "./tr[" + Convert.ToString(j) + "]/td";
                        title = docStockContext.DocumentNode.SelectNodes(path + "[2]//a")[0].InnerHtml.Trim();
                        title = title.Replace("'", "’");
                        department = docStockContext.DocumentNode.SelectNodes(path + "[3]/span/font")[0].InnerText.Trim();
                        time = docStockContext.DocumentNode.SelectNodes(path + "[4]/span/font")[0].InnerText.Trim();
                        url = "http://webapp.yuntech.edu.tw/WebMSS/" + docStockContext.DocumentNode.SelectSingleNode(path + "[2]/a[@href]").Attributes["href"].Value.Trim();
                        //textBox1.Text += "Title: " + title + "\r\nDepartment: " + department + "\r\ntime: " + time + "\r\nurl: " + url + "\r\n";

                        sheet.Cells[j, 1] = title;
                        sheet.Cells[j, 2] = department;
                        sheet.Cells[j, 3] = time;
                        sheet.Cells[j, 4] = url;
                    }
                }

                else if(i <= 6)
                {
                    ms = new MemoryStream(client.DownloadData(html_path[i]));
                    doc.Load(ms, Encoding.UTF8);
                    docStockContext.LoadHtml(doc.DocumentNode.SelectSingleNode("/html[1]/body[1]/div[1]/table[1]/tr[5]/td[2]/table[1]").InnerHtml);
                    HtmlNode sum_tr = docStockContext.DocumentNode.SelectSingleNode("/tr[last()]");
                    int sum = Get_Sum(sum_tr, "共", "則");
                    int now = 0;

                    do
                    {
                        ms = new MemoryStream(client.DownloadData(html_path[i] + now.ToString()));
                        doc.Load(ms, Encoding.UTF8);
                        docStockContext.LoadHtml(doc.DocumentNode.SelectSingleNode("/html[1]/body[1]/div[1]/table[1]/tr[5]/td[2]/table[1]").InnerHtml);
                        nodes = docStockContext.DocumentNode.SelectNodes(".//tr");
                        for (int j = 1; j <= nodes.Count() - 5; j += 2)
                        {
                            path = ".//tr[" + Convert.ToString(j) + "]/td";
                            title = docStockContext.DocumentNode.SelectNodes(path + "[3]/a")[0].InnerHtml.Trim();
                            title = title.Replace("'", "’");
                            department = docStockContext.DocumentNode.SelectNodes(path + "[4]//a")[0].InnerText.Trim();
                            time = docStockContext.DocumentNode.SelectNodes(path + "[5]/div")[0].InnerText.Trim();
                            url = docStockContext.DocumentNode.SelectSingleNode(path + "[3]/a[@href]").Attributes["href"].Value.Trim();
                            //textBox1.Text += "Title: " + title + "\r\nDepartment: " + department + "\r\ntime: " + time + "\r\nurl: " + url + "\r\n";
                            now++;
                            sheet.Cells[now + 1, 1] = title;
                            sheet.Cells[now + 1, 2] = department;
                            sheet.Cells[now + 1, 3] = time;
                            sheet.Cells[now + 1, 4] = url;
                        }
                    } while (sum - now > 0);
                }
                else if(i == 7)
                {
                    ms = new MemoryStream(client.DownloadData(html_path[i]));
                    doc.Load(ms, Encoding.UTF8);
                    docStockContext.LoadHtml(doc.DocumentNode.SelectSingleNode("/html[1]/body[1]/table[4]/tr").InnerHtml);
                    HtmlNode sum_tr = docStockContext.DocumentNode.SelectSingleNode("/td[1]");
                    int sum = Get_Sum(sum_tr, "：", "則");
                    int page = 1;
                    int now = 0;
                    do
                    {
                        ms = new MemoryStream(client.DownloadData(html_path[i] + page.ToString()));
                        doc.Load(ms, Encoding.UTF8);
                        docStockContext.LoadHtml(doc.DocumentNode.SelectSingleNode("/html[1]/body[1]/table[3]").InnerHtml);
                        nodes = docStockContext.DocumentNode.SelectNodes(".//tr");
                        for (int j = 1; j <= nodes.Count(); j ++)
                        {
                            now++;
                            if (now > 300)
                            {
                                continue;
                            }
                            path = ".//tr[" + Convert.ToString(j) + "]/td";
                            title = docStockContext.DocumentNode.SelectNodes(path + "[2]/font[1]/a")[0].InnerHtml.Trim();
                            title = title.Replace("'", "’");
                            department = docStockContext.DocumentNode.SelectNodes(path + "[3]")[0].InnerText.Trim();
                            time = docStockContext.DocumentNode.SelectNodes(path + "[1]")[0].InnerText.Trim();
                            url = "http://libweb.yuntech.edu.tw/news" + docStockContext.DocumentNode.SelectSingleNode(path + "[2]/font[1]/a[@href]").Attributes["href"].Value.Trim().Substring(1);
                            //textBox1.Text += "Title: " + title + "\r\nDepartment: " + department + "\r\ntime: " + time + "\r\nurl: " + url + "\r\n";
                            
                            sheet.Cells[now + 1, 1] = title;
                            sheet.Cells[now + 1, 2] = department;
                            sheet.Cells[now + 1, 3] = time;
                            sheet.Cells[now + 1, 4] = url;
                        }
                        page++;
                    } while (now < 300 && now < sum);
                }
                else
                {
                    ms = new MemoryStream(client.DownloadData(html_path[i]));
                    doc.Load(ms, Encoding.UTF8);
                    docStockContext.LoadHtml(doc.DocumentNode.SelectSingleNode("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[1]").InnerHtml);
                    HtmlNode sum_tr = docStockContext.DocumentNode.SelectSingleNode("/span[1]");
                    int sum = Get_Sum(sum_tr, "共", "頁");
                    int page = 1;
                    int now = 0;
                    do
                    {
                        ms = new MemoryStream(client.DownloadData(html_path[i] + page.ToString()));
                        doc.Load(ms, Encoding.UTF8);
                        docStockContext.LoadHtml(doc.DocumentNode.SelectSingleNode("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[2]/table[1]").InnerHtml);
                        nodes = docStockContext.DocumentNode.SelectNodes(".//tr");
                        for (int j = 1; j <= nodes.Count(); j++)
                        {
                            now++;
                            if (now > 300)
                            {
                                continue;
                            }
                            path = ".//tr[" + Convert.ToString(j) + "]/td";
                            title = docStockContext.DocumentNode.SelectNodes(path + "[1]/a")[0].InnerHtml.Trim();
                            title = title.Replace("'", "’");
                            url = docStockContext.DocumentNode.SelectSingleNode(path + "[1]/a[@href]").Attributes["href"].Value.Trim().Substring(1);
                            //textBox1.Text += "Title: " + title + "\r\nDepartment: " + department + "\r\ntime: " + time + "\r\nurl: " + url + "\r\n";
                            sheet.Cells[now + 1, 1] = title;
                            sheet.Cells[now + 1, 2] = "N";
                            sheet.Cells[now + 1, 3] = "N";
                            sheet.Cells[now + 1, 4] = url;
                        }
                        page++;
                    } while (now < 300 && page < sum);
                }
                sheet.Cells.Columns.AutoFit();
            }
            book.SaveCopyAs(Excel_path);
            book.Close(0);
            App.Quit();
        }

        private void Get_Announcement_Click(object sender, EventArgs e)
        {
            textBox1.Text += "取得公告中......\r\n";
            Get_Announcement_func();
            textBox1.Text += "成功取得公告資料\r\n";
        }

        public void Upload_Announcement_func()
        {
            string Excel_path = @"E:\Web Crawler\" + "Announcement.xlsx";
            App = new Excel.Application();
            book = App.Workbooks.Open(Excel_path, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                                  Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
            for (int p = 1; p <= 9; p++)
            {

                Excel.Worksheet sheet = (Excel.Worksheet)book.Worksheets.get_Item(p);
                int rows_count = sheet.UsedRange.Cells.Rows.Count;
                Excel.Range rng = sheet.Cells.get_Range("A2", "D" + rows_count);


                for (int i = 2; i <= rows_count; i++)
                {
                    string dbHost = "localhost";//資料庫位址
                    string dbUser = "root";//資料庫使用者帳號
                    string dbPass = "";//資料庫使用者密碼
                    string dbName = "csv_db";//資料庫名稱

                    string connStr = "server=" + dbHost + ";uid=" + dbUser + ";pwd=" + dbPass + ";database=" + dbName + ";CharSet =" + "utf8";
                    MySqlConnection conn = new MySqlConnection(connStr);
                    MySqlCommand command = conn.CreateCommand();
                    conn.Open();
                    command.CommandText = "Insert into " + "ann" + p.ToString() + "(matter,unit,date,site) values('" + sheet.Cells[i, 1].Value2 + "','" + sheet.Cells[i, 2].Value2 + "','" + sheet.Cells[i, 3].Text + "','" + sheet.Cells[i, 4].Value2 + "')";
                    command.ExecuteNonQuery();
                    conn.Close();
                }
            }
            book.Close(0);
            App.Quit();
        }

        private void Upload_Announcement_Click(object sender, EventArgs e)
        {
            textBox1.Text += "上傳公告中......\r\n";
            Upload_Announcement_func();
            textBox1.Text += "成功上傳公告資料\r\n";
        }

        private void Auto_upload_Timer_Tick(object sender, EventArgs e)
        {
            string year_str = DateTime.Now.ToString("yyyy");
            string month_str = DateTime.Now.ToString("MM");
            string hour_str = DateTime.Now.ToString("HH");
            int year = Int32.Parse(year_str);
            int month = Int32.Parse(month_str);
            int year_end = year + 1;
            int month_end = month;
            int hour = Int32.Parse(hour_str);
            string dbHost = "localhost";//資料庫位址
            string dbUser = "root";//資料庫使用者帳號
            string dbPass = "";//資料庫使用者密碼
            string dbName = "csv_db";//資料庫名稱

            string connStr = "server=" + dbHost + ";uid=" + dbUser + ";pwd=" + dbPass + ";database=" + dbName;

            MySqlConnection conn = new MySqlConnection(connStr);
            MySqlCommand command = conn.CreateCommand();

            if (hour >= 8 && hour <= 17)
            {
                Get_Announcement_func();
                for (int i = 1; i <= 5; i++)
                {
                    Clear_Table_func("ann" + i.ToString());
                }
                Upload_Announcement_func();
                textBox1.Text += "已於 " + DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss") + " 成功上傳公告。\r\n";
                if (hour == 12)
                {
                    for (int y = year, m = month; y * 100 + m <= year_end * 100 + month_end; m++)
                    {
                        if (m > 12)
                        {
                            m = 1;
                            y++;
                        }
                        Get_Calendar_func(y, m);
                        string name = y.ToString() + "_" + m.ToString().PadLeft(2, '0');
                        string cmdStr = "SELECT COUNT(*) FROM information_schema.tables WHERE table_schema = 'csv_db' AND table_name = '" + name + "'";

                        using (conn = new MySqlConnection(connStr))
                        {
                            MySqlCommand cmd = new MySqlCommand(cmdStr, conn);
                            conn.Open();
                            MySqlDataReader reader = cmd.ExecuteReader();
                            while (reader.Read())
                            {
                                int count = reader.GetInt32(0);
                                if (count == 0)
                                {
                                    // MessageBox.Show("No such data table exists!");
                                    string strData = System.IO.File.ReadAllText(@"D:\wamp\www\Read_template.php");
                                    strData = strData.Replace("file_name", name);
                                    System.IO.File.WriteAllText(@"D:\wamp\www\Read_" + name + ".php", strData);
                                    Create_Table_func(name);
                                }
                                else if (count == 1)
                                {
                                    // MessageBox.Show("Such data table exists!");
                                    Clear_Table_func(name);
                                }
                            }
                            Upload_Calendar_func(name);
                            conn.Close();
                        }
                    }
                    textBox1.Text += "已於 " + DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss") + " 成功上傳行事曆。\r\n";
                }
            }
        }

        private void Auto_Upload_ckbox_CheckedChanged(object sender, EventArgs e)
        {
            if (Auto_Upload_ckbox.Checked)
            {
                Auto_upload_Timer.Enabled = true;
            }

            else
            {
                Auto_upload_Timer.Enabled = false;
            }
        }

        private void EXIT_Click(object sender, EventArgs e)
        {
            foreach (System.Diagnostics.Process excelProc in System.Diagnostics.Process.GetProcessesByName("EXCEL"))
            {
                excelProc.Kill();
            }
            this.Close();
            Environment.Exit(Environment.ExitCode);
        }

        private void bt_TEST_Click(object sender, EventArgs e)
        {
            string year_str = DateTime.Now.ToString("yyyy");
            string month_str = DateTime.Now.ToString("MM");
            string hour_str = DateTime.Now.ToString("HH");
            int year = Int32.Parse(year_str);
            int month = Int32.Parse(month_str);
            int year_end = year + 1;
            int month_end = month;
            int hour = Int32.Parse(hour_str);
            string dbHost = "localhost";//資料庫位址
            string dbUser = "root";//資料庫使用者帳號
            string dbPass = "";//資料庫使用者密碼
            string dbName = "csv_db";//資料庫名稱

            string connStr = "server=" + dbHost + ";uid=" + dbUser + ";pwd=" + dbPass + ";database=" + dbName;
            for (int y = year, m = month; y * 100 + m <= year_end * 100 + month_end; m++)
            {
                if (m > 12)
                {
                    m = 1;
                    y++;
                }
                string name = y.ToString() + "_" + m.ToString().PadLeft(2, '0');
                string cmdStr = "SELECT COUNT(*) FROM information_schema.tables WHERE table_schema = 'csv_db' AND table_name = '" + name + "'";

                // MessageBox.Show("No such data table exists!");
                string strData = System.IO.File.ReadAllText(@"D:\wamp\www\Read_template.php");
                strData = strData.Replace("file_name", name);
                System.IO.File.WriteAllText(@"D:\wamp\www\Read_" + name + ".php", strData);
                //Create_Table_func(name);
            }
        }
    }
}
