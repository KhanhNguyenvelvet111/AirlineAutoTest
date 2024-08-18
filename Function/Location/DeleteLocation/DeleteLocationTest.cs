using Microsoft.VisualStudio.TestPlatform.ObjectModel;
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

namespace ProjectExcelReader.Function.Location.DeleteLocation
{
    public class DeleteLocationTest
    {
        IWebDriver driver;
        public DeleteLocationTest(IWebDriver web)
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

                for (int row = 17; row <= rowCount; row++)
                {
                    if (!worksheet.Cells[row, 2].Value.Equals("Xóa địa điểm"))
                    {
                        Console.WriteLine("Next case: " + worksheet.Cells[row, 2].Value);
                        break;
                    }

                    LocationData item = new LocationData();
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
            driver.FindElement(By.XPath("(//button[contains(text(),'delete')])[2]")).Click();
            Thread.Sleep(1000);
            driver.FindElement(By.CssSelector($".{data.actionType}")).Click();
            Thread.Sleep(1000);
            
            try
            {
                IWebElement item = driver.FindElement(By.XPath("//p[normalize-space()='Xóa Thành công']"));
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
