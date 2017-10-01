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
using System.Text;
using System.Threading.Tasks;

namespace DIA_HOM_NB
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
                    driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);
                    return driver;
                }


                else if (browser == "firefox")
                {
                    driver = new FirefoxDriver();
                   // driver.Manage().Window.Maximize();
                    driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(30);
                    return driver;
                }
                else if (browser == "IE")
                {
                    driver = new InternetExplorerDriver();
                    driver.Manage().Window.Maximize();
                    driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(30);
                    return driver;
                }
                else return null;
            }

            public static void waitTillElementToAppear(string element)
            {


                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(30));
                wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(element)));


            }



            public static void EnterText(string element, string value)
            {

                driver.FindElement(By.XPath(element)).Clear();
                driver.FindElement(By.XPath(element)).SendKeys(value);

            }

            public static void ClickPerform(string element)
            {
                IWebElement ele = driver.FindElement(By.XPath(element));

                Actions actions = new Actions(driver);

                actions.MoveToElement(ele).Click().Perform();



            }



            public static void EnterTextMove(string element, string value)
            {
                IWebElement ele = driver.FindElement(By.XPath(element));

                Actions actions = new Actions(driver);

                actions.MoveToElement(ele).Click();
                ele.SendKeys(value);



            }
            public static void SelectDropDown(string element, string value)
            {
                //if (type == Property_type.Id)
                //new SelectElement(Property_Collection.driver.FindElement(By.Id(element))).SelectByText(value);
                //if (type == Property_type.XPath)
                new SelectElement(driver.FindElement(By.XPath(element))).SelectByText(value);

            }

            public static void SelectTextDropDown(string element, string value)
            {
                if (value != "")
                {
                    driver.FindElement(By.XPath(element)).Clear();
                    driver.FindElement(By.XPath(element)).SendKeys(value);
                    driver.FindElement(By.XPath(element)).Click();

                }

               

            }

            public static void EnterTextTab(string element, string value)
            {
                if (value != "")
                {

                    driver.FindElement(By.XPath(element)).SendKeys(value);
                    System.Threading.Thread.Sleep(200);
                    driver.FindElement(By.XPath(element)).SendKeys(Keys.Tab);

                }

                

            }

            public static void EnterTextFocus(string element, string value)
        { 
                if (value != "")
                {


                    driver.FindElement(By.XPath(element)).SendKeys(Keys.Control + "a");
                    driver.FindElement(By.XPath(element)).SendKeys(value);
                    System.Threading.Thread.Sleep(200);
                    driver.FindElement(By.XPath(element)).SendKeys(Keys.Tab);


                }

               

            }

            public static void SelectListElement(string element, string value)
            {
                // driver.FindElement(By.XPath(element)).Clear();
                driver.FindElement(By.XPath(element)).SendKeys(value);
                IList<IWebElement> elements_to_be_click = driver.FindElements(By.XPath(element));
                foreach (IWebElement ele in elements_to_be_click)
                {

                    if (ele.TagName.Equals(value))
                    {
                        ele.Click();
                    }

                }


            }

            public static void Click(string element)
            {
            try
            {
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
                wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath(element)));
                driver.FindElement(By.XPath(element)).Click();
            }
            catch
            {
                driver.FindElement(By.XPath(element)).Click();
            }

            }
            public static void mouseHover(string element)
            {
                Actions hover = new Actions(driver);
                IWebElement topic = driver.FindElement(By.XPath(element));
                hover.MoveToElement(topic).Build().Perform();
            }

            public static void mouseclick(string element)
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
            public static void SelectTableEle(String element, String value)
            {
                var dropdown = Perform.driver.FindElement(By.XPath(element));
                new SelectElement(driver.FindElement(By.XPath(element))).SelectByText(value);
            }
            /*  public static void click_Checkbox(String element,String value)
              {

                  if (value=="YES")
                  Property_Collection.driver.FindElement(By.XPath(element)).Click();



              }*/

            public static string GetText(string element)
            {
               
                return driver.FindElement(By.XPath(element)).GetAttribute("value");

            }
            public static string GetTextFromDDL(string element, String type)
            {
                if (type == "Id")
                    return new SelectElement(driver.FindElement(By.Id(element))).AllSelectedOptions.SingleOrDefault().Text;
                if (type == "XPath")
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
                driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(30);
            }

            public static bool retryingFindClick(String element)
            {
                bool result = false;
                int attempts = 0;
                while (attempts < 2)
                {
                    try
                    {
                        driver.FindElement(By.XPath(element)).Click();
                        result = true;
                        break;
                    }
                    catch (Exception e)
                    {
                    }
                    attempts++;
                }
                return result;
            }

            public static void ScreenShot(string location)
            {
                ITakesScreenshot ss = driver as ITakesScreenshot;
                Screenshot screenshot = ss.GetScreenshot();
              
                screenshot.SaveAsFile(location, ScreenshotImageFormat.Png);
            }
        
        public static bool IsElementDisplayed(string element)
        {
            IReadOnlyCollection<IWebElement> elements = driver.FindElements(By.XPath(element));
            if (elements.Count > 0)
            {
                return elements.ElementAt(0).Displayed;
            }
            return false;
        }
        public static void CheckTitle(String ExpectedTitle)
        {
            if (driver.Title.Equals(ExpectedTitle))
            {
                test.Log(RelevantCodes.ExtentReports.LogStatus.Pass, "Page title is as expected :- " + ExpectedTitle);
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
