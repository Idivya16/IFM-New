using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.IE;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using RelevantCodes.ExtentReports;
using System;
using System.Collections.Generic;
using System.Linq;


namespace VR_COM_BOP
{
    class Perform
    {
       
        public static IWebDriver driver;
        public static ExtentReports report;
        public static ExtentTest test;
        public static IWebDriver Browser(String browser)
        {



            if (browser == "chrome")
            {
                driver = new ChromeDriver();
                driver.Manage().Window.Maximize();
               driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(15);
                return driver;
            }


            else if (browser == "firefox")
            {
                driver = new FirefoxDriver();
               // driver.Manage().Window.Maximize();
                driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(15);
                return driver;
            }
            else if (browser == "IE")
            {
                driver = new InternetExplorerDriver();
              // driver.Manage().Window.Maximize();
                driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(15);
                return driver;
            }
            else return null;
        }

        public static void waitTillElementToAppear(string element)
        {


            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(60));
            wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(element)));


        }

        public static void EnterText(string element, string value, Property_type type)
        {
            if (type == Property_type.Id)
            {
              driver.FindElement(By.Id(element)).Clear();
               driver.FindElement(By.Id(element)).SendKeys(value);
            }
            if (type == Property_type.XPath)
            {
                driver.FindElement(By.XPath(element)).Clear();
               driver.FindElement(By.XPath(element)).SendKeys(value);
            }
        }

        public static void SelectDropDown(string element, string value, Property_type type)
        {
            if (type == Property_type.Id)
                new SelectElement(driver.FindElement(By.Id(element))).SelectByText(value);
            if (type == Property_type.XPath)
                new SelectElement(Perform.driver.FindElement(By.XPath(element))).SelectByText(value); 
           

        }

        public static void Click(string element, Property_type type)
        {
           /* if (type == Property_type.Id)
                driver.FindElement(By.Id(element)).Click();*/
            if (type == Property_type.XPath)
                driver.FindElement(By.XPath(element)).Click();
           /* if (type == Property_type.CssName)
               driver.FindElement(By.CssSelector(element)).Click();
            if (type == Property_type.LinkText)
               driver.FindElement(By.LinkText(element)).Click();*/
        }

        public static void mouseHover(string element, Property_type type)
        {
            Actions hover = new Actions(driver);
            IWebElement topic = driver.FindElement(By.XPath(element));
            hover.MoveToElement(topic).Build().Perform();
        }

        public static void mouseclick(string element, Property_type type)
        {
            Actions popupclick = new Actions(driver);
            IWebElement accept = driver.FindElement(By.XPath(element));
            popupclick.MoveToElement(accept).Click();
        }

        public static void click_on_webElements(String locatorvalue, IWebDriver driver)
        {
            IList<IWebElement> elements_to_be_click = driver.FindElements(By.XPath(locatorvalue));
            foreach (IWebElement ele in elements_to_be_click)
            {

                ele.Click();

            }


        }
        public static void selectbyvalue(String element,String value)
            {
            new SelectElement(driver.FindElement(By.XPath(element))).SelectByValue(value);
        }
       /* public static void click_Checkbox(String element,String value)
          {

              if (value=="")
              Property_Collection.driver.FindElement(By.XPath(element)).Click();
              


          }*/
        public static string GetText(string element, Property_type type)
        {
            if (type == Property_type.Id)
                return driver.FindElement(By.Id(element)).GetAttribute("value");
            if (type == Property_type.XPath)
                return driver.FindElement(By.XPath(element)).GetAttribute("value");
            else return String.Empty;
        }
        public static string GetTextFromDDL(string element, Property_type type)
        {
            if (type == Property_type.Id)
                return new SelectElement(driver.FindElement(By.Id(element))).AllSelectedOptions.SingleOrDefault().Text;
            if (type == Property_type.XPath)
                return new SelectElement(driver.FindElement(By.XPath(element))).AllSelectedOptions.SingleOrDefault().Text;
            else return String.Empty;
        }
        /* public static bool isElementPresent(IWebElement element)
         {
             if (element.isDisplayed())
             {
                 try
                 {
                     element.Click();
                     return true;
                 }
                 catch (Exception e)
                 {
                     Console.WriteLine("Unable to find element ");
                     return false;
                 }
             }
             else
             {

                Console.WriteLine("Element is not displaying on the page");
                 return false;
             }

         }*/


        /* public static bool Boolean(string text)
         {
             if (Property_Collection.driver.PageSource.Contains(text))
                 return true;
             else
                 return false;
         }*/

        public static void Wait()
        {
           driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(60);
        }

        public static void ScreenShot(string location)
        {
            ITakesScreenshot ss =driver as ITakesScreenshot;
            Screenshot screenshot = ss.GetScreenshot();
            // DateTime time = DateTime.Now;
            // string dateToday = "_date_" + time.ToString("yyyy-MM-dd") + "_time_" + time.ToString("HH-mm-ss");
            screenshot.SaveAsFile(location, ScreenshotImageFormat.Png);
        }

        public static bool isAlertPresent()
        {
            try
            {
                driver.SwitchTo().Alert(); 
                return true;
            }
            catch(NoAlertPresentException Ex)
            {
                return false;
            }
        }

        public static void CheckTitle(String ExpectedTitle)
        {
            if(driver.Title.Equals(ExpectedTitle))
            {
                test.Log(LogStatus.Pass, "Page title is as expected :- " + ExpectedTitle);  
            }
            else
            {
                test.Log(LogStatus.Fail, "Incorrect Page");
            }
        }

        public static void PageContains(String ExpectedText)
        {
            if (driver.PageSource.Contains(ExpectedText))
            {
                test.Log(LogStatus.Pass, "Page Contains text as expected " + ExpectedText);
            }
            else
            {
                test.Log(LogStatus.Fail, "Incorrect Text");
            }
        }

        public static void IsElementPresent(String element)
        {

            try
            {
                if (driver.FindElement(By.XPath(element)).Displayed == true)
                {
                    test.Log(LogStatus.Pass, "Element present " + element);
                }
            }
            
             catch
            {
                test.Log(LogStatus.Fail, "Element not present");
            }
           
        }

       

    }
}





















