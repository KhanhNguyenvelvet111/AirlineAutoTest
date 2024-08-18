using Microsoft.VisualStudio.TestTools.UnitTesting;
using NUnit.Framework;
using OfficeOpenXml;
using OpenQA.Selenium;
using ProjectExcelReader.common;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;

namespace ProjectExcelReader.Function.Location.SearchLocation
{
    public class SearchLocationTest
    {
        IWebDriver driver;

        public SearchLocationTest(IWebDriver web)
        {
            this.driver = web;
        }

        public static LocationData[] getListDataExcel()
        {
            List<LocationData> dataResult = new List<LocationData>();

            using (var package = new ExcelPackage(new System.IO.FileInfo(common1.file)))
            {
                var worksheet = package.Workbook.Worksheets["Function"];
                int rowCount = worksheet.Dimension.Rows;
                int columnCount = worksheet.Dimension.Columns;

                for (int row = 50; row < rowCount; row ++)
                {
                    if (!worksheet.Cells[row, 2].Value.Equals("Tìm kiếm địa điểm"))
                    {
                        Console.WriteLine("Next case: " + worksheet.Cells[row, 2].Value);
                        break;
                    }

                    LocationData item = new LocationData();
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

        public void setExcel(LocationData data)
        {
            using (var package = new ExcelPackage(new FileInfo(common1.file)))
            {

                ExcelWorksheet worksheet = package.Workbook.Worksheets["Function"];
                worksheet.Cells[data.row, data.column].Value = data.actual;
                package.Save();
            }
        }


        public void runCase(LocationData data)
        {

            if (data.search != "null")
                driver.FindElement(By.CssSelector("input[placeholder='Nhập Tên Địa Điểm']")).SendKeys(data.search);

            driver.FindElement(By.XPath("//div[@class='search']//*[name()='svg']")).Click();
            Thread.Sleep(1000);

            try
            {
                IWebElement item = driver.FindElement(By.CssSelector("td:nth-child(1)"));
                if (item.GetAttribute("innerText").Contains("Hội An"))
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
