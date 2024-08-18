using Microsoft.Testing.Platform.Extensions.Messages;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using NUnit.Framework.Internal.Execution;
using OfficeOpenXml;
using OpenQA.Selenium;
using ProjectExcelReader.common;
using System;
using System.Collections.Generic;
using System.Data.Common;
using System.IO;
using System.Security.Permissions;
using System.Threading;

namespace ProjectExcelReader.Function
{
    public class CreateLocationTest
    {
        IWebDriver driver;
        public CreateLocationTest(IWebDriver web)
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

                for (int row = 2; row < rowCount; row = row + 3)
                {
                    if (!worksheet.Cells[row, 2].Value.Equals("Tạo địa điểm"))
                    {
                        Console.WriteLine("Next case: " + worksheet.Cells[row, 2].Value);
                        break;
                    }

                    LocationData item = new LocationData();
                    item.column = 8;
                    item.row = row;

                    item.locationName = (string)worksheet.Cells[row, 6].Value;
                    item.countryName = (string)worksheet.Cells[row + 1, 6].Value;
                    item.image = (string)worksheet.Cells[row + 2, 6].Value;
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
                //Console.WriteLine(data.status + "");
                //Console.WriteLine(data.expected + "");
                //Console.WriteLine(data.column + "");
                //Console.WriteLine(data.row + "");
                package.Save();
            }
        }


        public void runCase(LocationData data)
        {

            if (data.locationName != "null")
                driver.FindElement(By.CssSelector(".name input")).SendKeys(data.locationName);

            if (data.countryName != "null")
                driver.FindElement(By.CssSelector(".contry input")).SendKeys(data.countryName);

            if (data.image != "null")
                driver.FindElement(By.Id("file")).SendKeys(data.image);

            driver.FindElement(By.CssSelector(".ADD")).Click();
            Thread.Sleep(1000);

            try
            {
                IWebElement item = driver.FindElement(By.XPath("//p[contains(text(),'Tạo thành công')]"));
                if (item.GetAttribute("innerText").Equals("Tạo thành công"))
                {
                    data.actual = "Thông báo tạo thành công";
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
