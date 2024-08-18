using NUnit.Framework;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace ProjectExcelReader.Function
{
    internal class Login
    {
        IWebDriver driver; 

        public Login(IWebDriver driver)
        {
            this.driver = driver;
        }

        public void SetUp()
        {
            driver.Navigate().GoToUrl("http://localhost:8080/#/acount/login");
            driver.FindElement(By.XPath("(//input[@placeholder='Enter email'])[1]")).SendKeys("khanhrv111@gmail.com");
            driver.FindElement(By.XPath("(//input[@id='password'])[1]")).SendKeys("Khanh@123");
            driver.FindElement(By.XPath("(//button[contains(text(),'Đăng nhập')])[1]")).Click();
            driver.Manage().Window.Maximize();
            Thread.Sleep(2000);
        }
    }
}
