using HtmlAgilityPack;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace stock_auto
{
    //爬網頁部分另外寫，以免以後網頁的html結構有變動
    class WebDownLoad
    {
        public List<string> web(string code)
        {
            string[] txt = new string[] { };

            try
            {
                WebClient url = new WebClient();
                //爬 股票資訊
                MemoryStream ms = new MemoryStream(url.DownloadData("http://tw.stock.yahoo.com/q/q?s=" + code));
                HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
                doc.Load(ms, Encoding.Default);
                HtmlAgilityPack.HtmlDocument hdc = new HtmlAgilityPack.HtmlDocument();

                //XPath 來解讀它 /html[1]/body[1]/center[1]/table[2]/tr[1]/td[1]/table[1] 
                hdc.LoadHtml(doc.DocumentNode.SelectSingleNode("/html[1]/body[1]/center[1]/table[2]/tr[1]/td[1]/table[1]").InnerHtml);

                HtmlNodeCollection htnode = hdc.DocumentNode.SelectNodes("./tr[1]/th");
                txt = hdc.DocumentNode.SelectSingleNode("./tr[2]").InnerText.Trim().Split('\n');


                //清除資料
                doc = null;
                hdc = null;
                url = null;
                ms.Close();
            }
            catch (Exception ex)
            {
                Console.WriteLine("ERROR! ex = "+ex);
                Console.Read();
            }

            return txt.ToList();
        }
            
    }
}
