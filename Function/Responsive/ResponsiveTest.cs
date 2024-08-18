using OfficeOpenXml;
using OpenQA.Selenium;
using ProjectExcelReader.common;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace ProjectExcelReader.Function.Responsive
{
    public class ResponsiveTest
    {
        IWebDriver driver;
        public ResponsiveTest(IWebDriver web)
        {
            this.driver = web;
        }

        public static ResponsiveData[] getListDataExcel()
        {
            List<ResponsiveData> dataResult = new List<ResponsiveData>();

            using (var package = new ExcelPackage(new System.IO.FileInfo(common1.file)))
            {
                var worksheet = package.Workbook.Worksheets["Function"];
                int rowCount = worksheet.Dimension.Rows;
                int columnCount = worksheet.Dimension.Columns;

                for (int row = 99; row < rowCount; row = row + 2)
                {
                    if (!worksheet.Cells[row, 2].Value.Equals("Kiểm tra thích ứng của giao diện"))
                    {
                        Console.WriteLine("Next case: " + worksheet.Cells[row, 2].Value);
                        break;
                    }

                    ResponsiveData item = new ResponsiveData();
                    item.column = 8;
                    item.row = row;

                    item.width  = worksheet.Cells[row, 6].Value.ToString();
                    item.height  = worksheet.Cells[row + 1, 6].Value.ToString();
                    item.expected = (string)worksheet.Cells[row, 7].Value;
                    item.actual = (string)worksheet.Cells[row, 8].Value;
                    item.status = (string)worksheet.Cells[row, 9].Value;


                    dataResult.Add(item);

                }
            }
            return dataResult.ToArray();
        }

        public void setExcel(ResponsiveData data)
        {
            using (var package = new ExcelPackage(new FileInfo(common1.file)))
            {

                ExcelWorksheet worksheet = package.Workbook.Worksheets["Function"];
                worksheet.Cells[data.row, data.column].Value = data.actual;
                package.Save();
            }
        }


        public void runCase(ResponsiveData data)
        {
            double w = 0.0;
            double h = 0.0;
            bool succes = double.TryParse(data.width, out w) && double.TryParse(data.height, out h);

            if (succes)
            {
                driver.Manage().Window.Size = new Size((int)w, (int)h);
            }
            //driver.Manage().Window.Maximize();

            IWebElement item = driver.FindElement(By.CssSelector(".crud"));
            Size sizeItem = item.Size;
            if (sizeItem.Height == 64)
            {
                data.actual = $"Chiều cao form crud là {sizeItem.Height}";
                setExcel(data);
            }
            else
            {
                data.actual = "Chiều cao form crud đã thay đổi";
                setExcel(data);
            }
            //Console.WriteLine(w);
            //Console.WriteLine(h);
            //Console.WriteLine(sizeItem.Width);
            //Console.WriteLine(sizeItem.Height);
            //Console.WriteLine(data.width);
            //Console.WriteLine(data.height);

        }
    }
}
