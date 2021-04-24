using HtmlAgilityPack;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace stock_excel
{
    //2020.0423
    /*
     https://dotblogs.com.tw/marsxie/2018/01/21/135908
     可以抓歷史股價

     html 歷史
     https://www.twse.com.tw/exchangeReport/STOCK_DAY?response=html&date=20200423&stockNo=0056

     csv 歷史
     http://www.twse.com.tw/exchangeReport/STOCK_DAY?response=csv&date=20200422&stockNo=0056
         
         */

    class stock_history
    {
        #region 網路上查到的方法，可以學習

        public static void DownloadByTwse(string STOCK_CODE, string date)
        {
            Console.WriteLine("更新股價：" + STOCK_CODE);

            string download_url = "http://www.twse.com.tw/exchangeReport/STOCK_DAY?response=csv&date=" + date + "&stockNo=" + STOCK_CODE;

            // 系統睡10秒, 避免快速呼叫而被證交所擋ip
            System.Threading.Thread.Sleep(10000);
            string downloadedData = "";
            using (WebClient wClient = new WebClient())
            {
                try
                {
                    downloadedData = wClient.DownloadString(download_url);
                }
                catch (WebException ex)
                {
                    Console.WriteLine("更新股價失敗：" + STOCK_CODE + " " + ex.Message);
                }
            }
            if (downloadedData.Trim().Length == 0)
            {
                return;
            }
            // 證交所一次是回應一整個月的資料
            CSVReader csv = new CSVReader();
            string[] lineStrs = downloadedData.Split('\n');
            for (int i = 0; i < lineStrs.Length; i++)
            {
                string strline = lineStrs[i];
                if (i == 0 || i == 1 || strline.Trim().Length == 0)
                {
                    continue;
                }
                if (strline.IndexOf("說明:") > -1 || strline.IndexOf("符號說明") > -1 || strline.IndexOf("當日統計資訊含一般") > -1 || strline.IndexOf("ETF證券代號") > -1)
                {
                    continue;
                }

                ArrayList result = new ArrayList();
                csv.ParseCSVData(result, strline);
                string[] datas = (string[])result.ToArray(typeof(string));

                //檢查資料內容
                if (Convert.ToInt32(datas[1].Replace(",", "")) == 0 || datas[3] == "--" || datas[4] == "--" || datas[5] == "--" || datas[6] == "--")
                {
                    continue;
                }

                string code = STOCK_CODE; //代號
                string datea = datas[0]; //日期
                string open_price = datas[3];//開盤價
                string high_price = datas[4]; //最高價
                string low_price = datas[5]; //最低價
                string close_price = datas[6]; //收盤價
                string volume = datas[1]; //成交股數

                // 以下應用請自行處理
            }
        }

        public class CSVReader
        {
            private Stream objStream;
            private StreamReader objReader;

            //add name space System.IO.Stream
            public CSVReader()
            {

            }
            public CSVReader(Stream filestream) : this(filestream, null) { }
            public CSVReader(StreamReader strReader)
            {
                this.objReader = strReader;
            }
            public CSVReader(Stream filestream, Encoding enc)
            {
                this.objStream = filestream;
                //check the Pass Stream whether it is readable or not
                if (!filestream.CanRead)
                {
                    return;
                }
                objReader = (enc != null) ? new StreamReader(filestream, enc) : new StreamReader(filestream);
            }
            //parse the Line
            public string[] GetCSVLine()
            {
                string data = objReader.ReadLine();
                if (data == null) return null;
                if (data.Length == 0) return new string[0];
                //System.Collection.Generic
                ArrayList result = new ArrayList();
                //parsing CSV Data
                ParseCSVData(result, data);
                return (string[])result.ToArray(typeof(string));
            }

            public void ParseCSVData(ArrayList result, string data)
            {
                int position = -1;
                while (position < data.Length)
                    result.Add(ParseCSVField(ref data, ref position));
            }

            private string ParseCSVField(ref string data, ref int StartSeperatorPos)
            {
                if (StartSeperatorPos == data.Length - 1)
                {
                    StartSeperatorPos++;
                    return "";
                }

                int fromPos = StartSeperatorPos + 1;
                if (data[fromPos] == '"')
                {
                    int nextSingleQuote = GetSingleQuote(data, fromPos + 1);
                    int lines = 1;
                    while (nextSingleQuote == -1)
                    {
                        data = data + "\n" + objReader.ReadLine();
                        nextSingleQuote = GetSingleQuote(data, fromPos + 1);
                        lines++;
                        if (lines > 20)
                            throw new Exception("lines overflow: " + data);
                    }
                    StartSeperatorPos = nextSingleQuote + 1;
                    string tempString = data.Substring(fromPos + 1, nextSingleQuote - fromPos - 1);
                    tempString = tempString.Replace("'", "''");
                    return tempString.Replace("\"\"", "\"");
                }

                int nextComma = data.IndexOf(',', fromPos);
                if (nextComma == -1)
                {
                    StartSeperatorPos = data.Length;
                    return data.Substring(fromPos);
                }
                else
                {
                    StartSeperatorPos = nextComma;
                    return data.Substring(fromPos, nextComma - fromPos);
                }
            }

            private int GetSingleQuote(string data, int SFrom)
            {
                int i = SFrom - 1;
                while (++i < data.Length)
                    if (data[i] == '"')
                    {
                        if (i < data.Length - 1 && data[i + 1] == '"')
                        {
                            i++;
                            continue;
                        }
                        else
                            return i;
                    }
                return -1;
            }
        }
        #endregion

        #region 我的方法，查當月歷史股價
        public static void DownLoadStock(string stockID, string date)
        {
            string url = "https://www.twse.com.tw/exchangeReport/STOCK_DAY?response=html&date=" + date + "&stockNo=" + stockID;

            WebClient webUrl = new WebClient();
            //爬 股票資訊
            MemoryStream ms  = new MemoryStream(webUrl.DownloadData(url));

            HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
            doc.Load(ms, Encoding.UTF8);//不用UTF8的話會是亂碼

            #region 抓出表格裡面的資訊

            List<stock_Data> stock_Datas       = new List<stock_Data>();


            HtmlAgilityPack.HtmlDocument hbody = new HtmlAgilityPack.HtmlDocument();
            hbody.LoadHtml(doc.DocumentNode.SelectSingleNode("/html/body/div/table/tbody").InnerHtml);
            //先抓出tr總共有幾筆
            HtmlNodeCollection tr = hbody.DocumentNode.SelectNodes("./tr");

            for (int i = 1; i < tr.Count; i++)
            {
                try
                {
                    hbody.LoadHtml(doc.DocumentNode.SelectSingleNode("/html/body/div/table/tbody/tr[" + i + "]").InnerHtml);
                    if (hbody != null)
                    {
                        stock_Data stock_Data = new stock_Data();

                        //表格裡的資料  
                        stock_Data.Date           = hbody.DocumentNode.SelectSingleNode("./td[1]").InnerText.Trim();
                        stock_Data.Deal_Stock_Num = hbody.DocumentNode.SelectSingleNode("./td[2]").InnerText.Trim();
                        stock_Data.Deal_Price     = hbody.DocumentNode.SelectSingleNode("./td[3]").InnerText.Trim();
                        stock_Data.Start_Price    = hbody.DocumentNode.SelectSingleNode("./td[4]").InnerText.Trim();
                        stock_Data.High_Price     = hbody.DocumentNode.SelectSingleNode("./td[5]").InnerText.Trim();
                        stock_Data.Low_Price      = hbody.DocumentNode.SelectSingleNode("./td[6]").InnerText.Trim();
                        stock_Data.Close_Price    = hbody.DocumentNode.SelectSingleNode("./td[7]").InnerText.Trim();
                        stock_Data.UpDownPrice    = hbody.DocumentNode.SelectSingleNode("./td[8]").InnerText.Trim();
                        stock_Data.Deal_Num       = hbody.DocumentNode.SelectSingleNode("./td[9]").InnerText.Trim();


                        stock_Datas.Add(stock_Data);
                    }
                    else
                    {
                        return;
                    }

                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message + "\r\n");
                }

            }

            //資料都已塞入stock_Datas，看要塞資料庫還怎樣~



            #endregion
        }

        public class stock_Data
        {
            public string Date           { get; set; }//日期
            public string Deal_Stock_Num { get; set; }//成交股數
            public string Deal_Price     { get; set; }//成交金額
            public string Start_Price    { get; set; }//開盤價
            public string High_Price     { get; set; }//最高價
            public string Low_Price      { get; set; }//最低價
            public string Close_Price    { get; set; }//收盤價
            public string UpDownPrice    { get; set; }//漲跌價差
            public string Deal_Num       { get; set; }//成交筆數
        }
        #endregion
    }
}
