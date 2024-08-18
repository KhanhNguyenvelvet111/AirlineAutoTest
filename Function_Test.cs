using Microsoft.VisualStudio.TestTools.UnitTesting;
using NUnit.Framework;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium;
using System;
using System.Threading;
using ProjectExcelReader.Function;
using ProjectExcelReader.Function.Location.DeleteLocation;
using ProjectExcelReader.Function.Location.SearchLocation;
using ProjectExcelReader.Function.TradeMark;
using ProjectExcelReader.Function.TradeMark.CreateTradeMark;
using ProjectExcelReader.Function.TradeMark.SearchTradeMark;
using ProjectExcelReader.Function.Location.UpdateLocation;
using ProjectExcelReader.Function.TradeMark.DeteleTradeMark;
using ProjectExcelReader.Function.TradeMark.UpdateTradeMark;
using ProjectExcelReader.Function.Responsive;

namespace ProjectExcelReader
{
    [TestClass]
    public class Function_Test
    {
        IWebDriver driver;
        [SetUp]
        public void SetUp()
        {
            driver = new ChromeDriver();

        }

        //Location

        [TestCaseSource(typeof(CreateLocationTest), nameof(CreateLocationTest.getListDataExcel))]
        public void createLocation(LocationData data)
        {
            Login datas = new Login(driver);
            datas.SetUp();

            driver.Navigate().GoToUrl("http://localhost:8080/#/admin/locations");
            CreateLocationTest u = new CreateLocationTest(driver);
            u.runCase(data);
            NUnit.Framework.Assert.That(data.actual, Is.Not.Null);
        }

        [TestCaseSource(typeof(DeleteLocationTest), nameof(DeleteLocationTest.getListDataExcel))]
        public void deleteLocation(LocationData data)
        {
            Login datas = new Login(driver);
            datas.SetUp();

            driver.Navigate().GoToUrl("http://localhost:8080/#/admin/locations");
            DeleteLocationTest u = new DeleteLocationTest(driver);
            u.runCase(data);
            NUnit.Framework.Assert.That(data.actual, Is.Not.Null);
        }

        [TestCaseSource(typeof(SearchLocationTest), nameof(SearchLocationTest.getListDataExcel))]
        public void searchLocation(LocationData data)
        {
            Login datas = new Login(driver);
            datas.SetUp();
            driver.Navigate().GoToUrl("http://localhost:8080/#/admin/locations");
            SearchLocationTest u = new SearchLocationTest(driver);
            u.runCase(data);
            NUnit.Framework.Assert.That(data.actual, Is.Not.Null);
        }

        [TestCaseSource(typeof(UpdateLocationTest), nameof(UpdateLocationTest.getListDataExcel))]
        public void updateLocation(LocationData data)
        {
            Login datas = new Login(driver);
            datas.SetUp();
            driver.Navigate().GoToUrl("http://localhost:8080/#/admin/locations");
            UpdateLocationTest u = new UpdateLocationTest(driver);
            u.runCase(data);
            NUnit.Framework.Assert.That(data.actual, Is.Not.Null);
        }

        //TradeMark

        [TestCaseSource(typeof(CreateTradeMarkTest), nameof(CreateTradeMarkTest.getListDataExcel))]
        public void createTradeMark(TradeMarkData data)
        {
            Login datas = new Login(driver);
            datas.SetUp();
            driver.Navigate().GoToUrl("http://localhost:8080/#/admin/trademark");
            CreateTradeMarkTest u = new CreateTradeMarkTest(driver);
            u.runCase(data);
            NUnit.Framework.Assert.That(data.actual, Is.Not.Null);
        }

        [TestCaseSource(typeof(SearchTradeMarkTest), nameof(SearchTradeMarkTest.getListDataExcel))]
        public void searchTradeMark(TradeMarkData data)
        {
            Login datas = new Login(driver);
            datas.SetUp();
            driver.Navigate().GoToUrl("http://localhost:8080/#/admin/trademark");
            SearchTradeMarkTest u = new SearchTradeMarkTest(driver);
            u.runCase(data);
            NUnit.Framework.Assert.That(data.actual, Is.Not.Null);
        }

        [TestCaseSource(typeof(DeleteTradeMarkTest), nameof(DeleteTradeMarkTest.getListDataExcel))]
        public void deleteTradeMark(TradeMarkData data)
        {
            Login datas = new Login(driver);
            datas.SetUp();
            driver.Navigate().GoToUrl("http://localhost:8080/#/admin/trademark");
            DeleteTradeMarkTest u = new DeleteTradeMarkTest(driver);
            u.runCase(data);
            NUnit.Framework.Assert.That(data.actual, Is.Not.Null);
        }

        [TestCaseSource(typeof(UpdateTradeMarkTest), nameof(UpdateTradeMarkTest.getListDataExcel))]
        public void updateTradeMark(TradeMarkData data)
        {
            Login datas = new Login(driver);
            datas.SetUp();
            driver.Navigate().GoToUrl("http://localhost:8080/#/admin/trademark");
            UpdateTradeMarkTest u = new UpdateTradeMarkTest(driver);
            u.runCase(data);
            NUnit.Framework.Assert.That(data.actual, Is.Not.Null);
        }


        //Responsive

        [TestCaseSource(typeof(ResponsiveTest), nameof(ResponsiveTest.getListDataExcel))]
        public void responsive(ResponsiveData data)
        {
            Login datas = new Login(driver);
            datas.SetUp();
            driver.Navigate().GoToUrl("http://localhost:8080/#/admin/trademark");
            ResponsiveTest u = new ResponsiveTest(driver);
            u.runCase(data);
            NUnit.Framework.Assert.That(data.actual, Is.Not.Null);
        }

        [TearDown]
        public void TearDown()
        {
            Thread.Sleep(5000);
            driver.Quit();
        }

    }
}
