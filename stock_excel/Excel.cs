using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace stock_excel
{
    class Excel
    {
        public void Write()
        {
            string Folder = ConfigurationManager.AppSettings["ExcelFileSavePath"] ?? string.Empty;

            //先檢查目錄在不在 不在->建立目錄
            if (!Directory.Exists(Folder))
            {
                Directory.CreateDirectory(Folder);
            }

            #region 沒寫會出錯，不知道為什麼
            // If you are a commercial business and have
            // purchased commercial licenses use the static property
            // LicenseContext of the ExcelPackage class :
            ExcelPackage.LicenseContext = LicenseContext.Commercial;

            // If you use EPPlus in a noncommercial context
            // according to the Polyform Noncommercial license:
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            #endregion

            try
            {
                //在記憶體中建立一個Excel物件
                using (ExcelPackage p = new ExcelPackage())
                {
                    //第一張sheet 名字 查詢條件   //加入一個Sheet
                    ExcelWorksheet sheet = p.Workbook.Worksheets.Add("Stock");
                    //取得剛剛加入的Sheet(實體Sheet就叫MySheet)
                    ExcelWorksheet sheet1 = p.Workbook.Worksheets["Stock"];//取得Sheet1 

                    //標題
                    sheet1.Cells[1, 1].Value = "代號";
                    sheet1.Cells[1, 2].Value = "名字";
                    sheet1.Cells[1, 3].Value = "成交";
                    sheet1.Cells[1, 4].Value = "買進";
                    sheet1.Cells[1, 5].Value = "賣出";
                    sheet1.Cells[1, 6].Value = "漲跌";
                    sheet1.Cells[1, 7].Value = "張數";
                    sheet1.Cells[1, 8].Value = "昨收";
                    sheet1.Cells[1, 9].Value = "開盤";
                    sheet1.Cells[1, 10].Value = "最高";
                    sheet1.Cells[1, 11].Value = "最低";
                    sheet1.Cells[1, 12].Value = "建立時間";




                    for (int i = 0; i < stockDetail.Count; i++)
                    {
                        sheet1.Cells[i + 2, 1].Value = stockDetail[i].Code;
                        sheet1.Cells[i + 2, 2].Value = stockDetail[i].Name;
                        sheet1.Cells[i + 2, 3].Value = stockDetail[i].Deal;
                        sheet1.Cells[i + 2, 4].Value = stockDetail[i].Buy;
                        sheet1.Cells[i + 2, 5].Value = stockDetail[i].Sell;
                        sheet1.Cells[i + 2, 6].Value = stockDetail[i].Up_Down;
                        sheet1.Cells[i + 2, 7].Value = stockDetail[i].Number;
                        sheet1.Cells[i + 2, 8].Value = stockDetail[i].Receive;
                        sheet1.Cells[i + 2, 9].Value = stockDetail[i].Start;
                        sheet1.Cells[i + 2, 10].Value = stockDetail[i].High;
                        sheet1.Cells[i + 2, 11].Value = stockDetail[i].Low;
                        sheet1.Cells[i + 2, 12].Value = DateTime.Now.ToString("yyyy/MM/dd hh:mm:ss");
                    }


                    //設定自動適應寬度
                    sheet1.Cells.AutoFitColumns();

                    #region 存檔
                    //設定路徑
                    string localFilePath = Folder + @"\" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx";

                    if (!System.IO.Directory.Exists(Folder))
                    {
                        using (FileStream createStream = new FileStream(localFilePath, FileMode.Create, FileAccess.Write, FileShare.ReadWrite))
                        {
                            p.SaveAs(createStream);//存檔
                        }
                    }
                    else
                    {
                        using (FileStream createStream = new FileStream(localFilePath, FileMode.Create, FileAccess.Write, FileShare.ReadWrite))
                        {
                            p.SaveAs(createStream);//存檔
                        }

                    }
                    #endregion

                }
            }
            catch (Exception ex)
            {
                Console.Write("ex=" + ex);
                Console.WriteLine();
                Console.Read();//停下來
            }

        }
    }
}
