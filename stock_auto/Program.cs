using HtmlAgilityPack;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using System.Threading;
using System.Net.Mail;

namespace stock_auto
{
    //2020.0407 無聊玩一下爬蟲
    class Program
    {   //個股各日成交歷史 - html
        //https://www.twse.com.tw/exchangeReport/STOCK_DAY?response=html&date=20200423&stockNo=0056

        //個股各日成交歷史 - csv
        //http://www.twse.com.tw/exchangeReport/STOCK_DAY?response=csv&date=20200422&stockNo=0056

        static void Main(string[] args)
        {
            string connString = ConfigurationManager.ConnectionStrings["test"].ConnectionString;
            string stock_code = ConfigurationManager.AppSettings["stock_code"];

            //喜歡的股票代號由app.config決定
            List<string> stockList   = stock_code.Split(',').ToList();

            //喜歡的股票代號
            //List<string> stockList   = new List<string> {"0056", "0050", "2886", "2884", "2330" };
            List<Stock> stockDetail  = new List<Stock>();

            for (int i = 0; i < stockList.Count; i++)
            {
                //休息10毫秒
                Thread.Sleep(10);

                WebDownLoad webDownLoad = new WebDownLoad();
                List<string> list =  webDownLoad.web(stockList[i]);              

                
                Stock stock         = new Stock();
                //取出資訊
                string Code         = list[0].Trim().Replace("加到投資組合", "").ToString();  //股票代號

                int     CodeLength  = 0;
                CodeLength          = stockList[i].Length;

                stock.Code        = Code.Substring(0, CodeLength);//切出代碼
                stock.Name        = Code.Substring(4);          //切出股票的名字

                stock.Deal        = list[2].Trim().ToString();  //成交
                stock.Buy         = list[3].Trim().ToString();  //買進
                stock.Sell        = list[4].Trim().ToString();  //賣出
                stock.Up_Down     = list[5].Trim().ToString();  //漲跌
                stock.Number      = list[6].Trim().ToString();  //張數
                stock.Receive     = list[7].Trim().ToString();  //昨收
                stock.Start       = list[8].Trim().ToString();  //開盤
                stock.High        = list[9].Trim().ToString();  //最高
                stock.Low         = list[10].Trim().ToString(); //最低
                stockDetail.Add(stock);

                using (SqlConnection conn = new SqlConnection(connString))
                {
                    conn.Open();
                    using (SqlCommand cmd = new SqlCommand())
                    {
                        cmd.Connection = conn;
                        cmd.CommandText = @" 
                                             INSERT INTO [STOCK]
                                             ([Code], [Name], [Deal], [Buy], [Sell], [Up_Down], [Number], [Receive], [Start], [High], [Low], [CreateTime] )
                                             VALUES
                                             (@Code,  @Name,  @Deal,  @Buy,  @Sell,  @Up_Down,  @Number,  @Receive,  @Start,  @High,  @Low,  @CreateTime )
                                        ";

                        cmd.Parameters.Clear();
                        cmd.Parameters.AddWithValue("Code",       stock.Code);
                        cmd.Parameters.AddWithValue("Name",       stock.Name);
                        cmd.Parameters.AddWithValue("Deal",       stock.Deal);
                        cmd.Parameters.AddWithValue("Buy",        stock.Buy);
                        cmd.Parameters.AddWithValue("Sell",       stock.Sell);
                        cmd.Parameters.AddWithValue("Up_Down",    stock.Up_Down);
                        cmd.Parameters.AddWithValue("Number",     stock.Number);
                        cmd.Parameters.AddWithValue("Receive",    stock.Receive);
                        cmd.Parameters.AddWithValue("Start",      stock.Start);
                        cmd.Parameters.AddWithValue("High",       stock.High);
                        cmd.Parameters.AddWithValue("Low",        stock.Low);
                        cmd.Parameters.AddWithValue("CreateTime", Convert.ToDateTime(DateTime.Now.ToString()));

                        cmd.ExecuteNonQuery();

                    }
                }
                                

            }

            // 看要不要寫excel  1->寫  0-->不寫，外面決定
            if (ConfigurationManager.AppSettings["Excel"].Equals("1"))
            {
                Excel ex = new Excel();
                ex.Write(stockDetail);
            }

        }

        /// <summary>
        /// 有需要再用吧
        /// https://dotblogs.com.tw/chichiblog/2018/04/20/122816
        /// </summary>
        static void sendGmail()
        {
            MailMessage mail = new MailMessage();
            //前面是發信email後面是顯示的名稱
            mail.From = new MailAddress("mia550999@gmail.com", "信件名稱");

            //收信者email
            mail.To.Add("hulily8404@gmail.com");

            //設定優先權
            mail.Priority = MailPriority.Normal;

            //標題
            mail.Subject = "AutoEmail";

            //內容
            mail.Body = "<h1>HIHI,Wellcome</h1>";

            //內容使用html
            mail.IsBodyHtml = true;

            //設定gmail的smtp (這是google的)
            SmtpClient MySmtp = new SmtpClient("smtp.gmail.com", 587);

            //您在gmail的帳號密碼
            MySmtp.Credentials = new System.Net.NetworkCredential("account", "pw");

            //開啟ssl
            MySmtp.EnableSsl = true;

            //發送郵件
            MySmtp.Send(mail);

            //放掉宣告出來的MySmtp
            MySmtp = null;

            //放掉宣告出來的mail
            mail.Dispose();
        }

        public class Stock
        {
            public string Code    { get; set; }//股票代號
            public string Name    { get; set; }//股票名字
            public string Deal    { get; set; }//成交
            public string Buy     { get; set; }//買進
            public string Sell    { get; set; }//賣出
            public string Up_Down { get; set; }//漲跌
            public string Number  { get; set; }//張數
            public string Receive { get; set; }//昨收
            public string Start   { get; set; }//開盤
            public string High    { get; set; }//最高
            public string Low     { get; set; }//最低
        }
    }
}
