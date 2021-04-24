using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;

namespace stock_excel
{
    /*一天做一次，將資料庫的資料撈出，寫到excel*/
    class Program
    {
        static void Main(string[] args)
        {
            // 看要不要寫excel  1->寫  0-->不寫，外面決定
            if (ConfigurationManager.AppSettings["Excel"].Equals("1"))
            {
                Excel ex = new Excel();
                ex.Write();
            }

            //目前還沒用這隻
            stock_history.DownloadByTwse("0056", "20200423");

            //我的寫法
            stock_history.DownLoadStock("0056", "20200423");
        }
    }
}
