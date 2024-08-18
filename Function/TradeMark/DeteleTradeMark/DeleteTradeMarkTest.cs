using OfficeOpenXml;
using OpenQA.Selenium;
using ProjectExcelReader.common;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace ProjectExcelReader.Function.TradeMark.DeteleTradeMark
{
    public class DeleteTradeMarkTest
    {
        IWebDriver driver;
        public DeleteTradeMarkTest(IWebDriver web)
        {
            this.driver = web;
        }

        public static TradeMarkData[] getListDataExcel()
        {
            List<TradeMarkData> dataResult = new List<TradeMarkData>();

            using (var package = new ExcelPackage(new System.IO.FileInfo(common1.file)))
            {
                var worksheet = package.Workbook.Worksheets["Function"];

                int rowCount = worksheet.Dimension.Rows;
                int columnCount = worksheet.Dimension.Columns;

                for (int row = 70; row <= rowCount; row++)
                {
                    if (!worksheet.Cells[row, 2].Value.Equals("Xóa hãng máy bay"))
                    {
                        Console.WriteLine("Next case: " + worksheet.Cells[row, 2].Value);
                        break;
                    }

                    TradeMarkData item = new TradeMarkData();
                    item.column = 8;
                    item.row = row;
                    item.actionType = (string)worksheet.Cells[row, 6].Value;
                    item.key = (string)worksheet.Cells[row, 3].Value;
                    item.actual = (string)worksheet.Cells[row, 8].Value;
                    dataResult.Add(item);

                }
            }
            return dataResult.ToArray();
        }

        public void setExcel(TradeMarkData data)
        {
            using (var package = new ExcelPackage(new FileInfo(common1.file)))
            {

                ExcelWorksheet worksheet = package.Workbook.Worksheets["Function"];
                worksheet.Cells[data.row, data.column].Value = data.actual;
                package.Save();
            }
        }



        public void runCase(TradeMarkData data)
        {
            driver.FindElement(By.CssSelector("tr:nth-child(2) td:nth-child(3) button:nth-child(1)")).Click();
            Thread.Sleep(1000);
            driver.FindElement(By.CssSelector($".{data.actionType}")).Click();
            Thread.Sleep(1000);

            try
            {
                IWebElement item = driver.FindElement(By.XPath("(//p[normalize-space()='Xóa Thành công'])[1]"));
                if (item.GetAttribute("innerText").Equals("Xóa Thành công"))
                {
                    data.actual = "Thông báo xóa thành công";
                    setExcel(data);
                }
                else
                {
                    data.actual = "Hiển thị Error Message";
                    setExcel(data);
                }
            }
            catch (NoSuchElementException)
            {
                data.actual = "Hiển thị Error Message";
                setExcel(data);
            }
        }
    }
}
