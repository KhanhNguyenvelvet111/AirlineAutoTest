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

namespace ProjectExcelReader.Function.TradeMark.SearchTradeMark
{
    public class SearchTradeMarkTest
    {
        IWebDriver driver;

        public SearchTradeMarkTest(IWebDriver web)
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

                for (int row = 79; row < rowCount; row++)
                {
                    if (!worksheet.Cells[row, 2].Value.Equals("Tìm kiếm hãng máy bay"))
                    {
                        Console.WriteLine("Next case: " + worksheet.Cells[row, 2].Value);
                        break;
                    }

                    TradeMarkData item = new TradeMarkData();
                    item.column = 8;
                    item.row = row;

                    item.search = (string)worksheet.Cells[row, 6].Value;
                    item.key = (string)worksheet.Cells[row, 3].Value;
                    item.expected = (string)worksheet.Cells[row, 7].Value;
                    item.actual = (string)worksheet.Cells[row, 8].Value;
                    item.status = (string)worksheet.Cells[row, 9].Value;


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
                Console.WriteLine(data.search + "");
                Console.WriteLine(data.expected + "");
                Console.WriteLine(data.column + "");
                Console.WriteLine(data.row + "");
                package.Save();
            }
        }


        public void runCase(TradeMarkData data)
        {

            if (data.search != "null")
                driver.FindElement(By.XPath("(//input[@placeholder='Nhập Tên hãng bay'])[1]")).SendKeys(data.search);

            driver.FindElement(By.XPath("(//*[name()='svg'][@class='svg-inline--fa fa-magnifying-glass icon'])[1]")).Click();
            Thread.Sleep(1000);

            try
            {
                IWebElement item = driver.FindElement(By.CssSelector("td:nth-child(1)"));
                if (item.GetAttribute("innerText").Contains("Delta"))
                {
                    data.actual = "Hiển thị danh sách tìm";
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
