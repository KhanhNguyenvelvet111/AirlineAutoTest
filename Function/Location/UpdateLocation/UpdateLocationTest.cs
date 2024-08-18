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

namespace ProjectExcelReader.Function.Location.UpdateLocation
{
    public class UpdateLocationTest
    {
        IWebDriver driver;
        public UpdateLocationTest(IWebDriver web)
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

                for (int row = 20; row < rowCount; row = row + 4)
                {
                    if (!worksheet.Cells[row, 2].Value.Equals("Cập nhật địa điểm"))
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
                    item.actionType = (string)worksheet.Cells[row + 3, 6].Value;
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

            driver.FindElement(By.CssSelector("tr:nth-child(3) td:nth-child(1)")).Click();
            Thread.Sleep(1000);

            driver.FindElement(By.CssSelector(".name input")).SendKeys(Keys.End + Keys.Shift + Keys.Home + Keys.Delete);
            driver.FindElement(By.CssSelector(".contry input")).SendKeys(Keys.End + Keys.Shift + Keys.Home + Keys.Delete);

            if (data.locationName != "null")         
                driver.FindElement(By.CssSelector(".name input")).SendKeys(data.locationName);
            
            if (data.countryName != "null")
                driver.FindElement(By.CssSelector(".contry input")).SendKeys(data.countryName);
            
            if (data.image != "null")
                driver.FindElement(By.Id("file")).SendKeys(data.image);





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
