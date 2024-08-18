﻿using Microsoft.VisualStudio.TestTools.UnitTesting;
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

namespace ProjectExcelReader.Function.TradeMark.UpdateTradeMark
{
    public class UpdateTradeMarkTest
    {
        IWebDriver driver;
        public UpdateTradeMarkTest(IWebDriver web)
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

                for (int row = 73; row < rowCount; row = row + 3)
                {
                    if (!worksheet.Cells[row, 2].Value.Equals("Cập nhật hãng máy bay"))
                    {
                        Console.WriteLine("Next case: " + worksheet.Cells[row, 2].Value);
                        break;
                    }

                    TradeMarkData item = new TradeMarkData();
                    item.column = 8;
                    item.row = row;

                    item.tradeMarkName = (string)worksheet.Cells[row, 6].Value;
                    item.image = (string)worksheet.Cells[row + 1, 6].Value;
                    item.actionType = (string)worksheet.Cells[row + 2, 6].Value;

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
                package.Save();
            }
        }


        public void runCase(TradeMarkData data)
        {

            driver.FindElement(By.CssSelector("tr:nth-child(2) td:nth-child(1)")).Click();
            Thread.Sleep(1000);

            driver.FindElement(By.CssSelector("input[placeholder='Nhập...']")).SendKeys(Keys.End + Keys.Shift + Keys.Home + Keys.Delete);

            if (data.tradeMarkName != "null")
                driver.FindElement(By.CssSelector("input[placeholder='Nhập...']")).SendKeys(data.tradeMarkName);

            if (data.image != "null")
                driver.FindElement(By.XPath("(//input[@id='file'])[1]")).SendKeys(data.image);

            Thread.Sleep(1000);
            driver.FindElement(By.CssSelector($"{data.actionType}")).Click();
            Thread.Sleep(1000);

            try
            {
                IWebElement item = driver.FindElement(By.XPath("(//p[contains(text(),'Cập nhật thành công')])[1]"));
                if (item.GetAttribute("innerText").Equals("Cập nhật thành công"))
                {
                    data.actual = "Thông báo cập nhật thành công";
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
