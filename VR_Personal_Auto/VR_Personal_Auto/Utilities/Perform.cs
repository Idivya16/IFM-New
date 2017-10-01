using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.IE;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using RelevantCodes.ExtentReports;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace VR_Personal_Auto

{
    public class Perform
 

    {
        public static ExtentReports report;
        public static ExtentTest test;

        public static IWebDriver Browser(String browser)
            
        {



            if (browser == "chrome")
            {
                Property_Collection.driver = new ChromeDriver();
                Property_Collection.driver.Manage().Window.Maximize();
                Property_Collection.driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);
                return Property_Collection.driver;
            }


            else if (browser == "firefox")
            {
                Property_Collection.driver = new FirefoxDriver();
                Property_Collection.driver.Manage().Window.Maximize();
                Property_Collection.driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);
                return Property_Collection.driver;
            }
            else if (browser == "IE")
            {
                Property_Collection.driver = new InternetExplorerDriver();
                Property_Collection.driver.Manage().Window.Maximize();
                Property_Collection.driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);
                return Property_Collection.driver;
            }
            else return null;
        }

        public static void waitTillElementToAppear(string element)
        {

           
                WebDriverWait wait = new WebDriverWait(Property_Collection.driver,TimeSpan.FromSeconds(20));
            wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(element)));


        }



        public static void EnterText(string element, string value, Property_type type)
        {
            if (type == Property_type.Id)
            {
                Property_Collection.driver.FindElement(By.Id(element)).Clear();
                Property_Collection.driver.FindElement(By.Id(element)).SendKeys(value);
            }
            if (type == Property_type.XPath)
            {
                Property_Collection.driver.FindElement(By.XPath(element)).Clear();
                Property_Collection.driver.FindElement(By.XPath(element)).SendKeys(value);
            }
        }

        public static void SelectDropDown(string element, string value, Property_type type)
        {
            if (type == Property_type.Id)
                new SelectElement(Property_Collection.driver.FindElement(By.Id(element))).SelectByText(value);
            if (type == Property_type.XPath)
                new SelectElement(Property_Collection.driver.FindElement(By.XPath(element))).SelectByText(value);
            
        }

        public static void Click(string element, Property_type type)
        {
            if (type == Property_type.Id)
                Property_Collection.driver.FindElement(By.Id(element)).Click();
            if (type == Property_type.XPath)
                Property_Collection.driver.FindElement(By.XPath(element)).Click();
            if (type == Property_type.CssName)
                Property_Collection.driver.FindElement(By.CssSelector(element)).Click();
            if (type == Property_type.LinkText)
                Property_Collection.driver.FindElement(By.LinkText(element)).Click();
        }
        public static void mouseHover(string element, Property_type type)
        {
            Actions hover = new Actions(Property_Collection.driver);
            IWebElement topic = Property_Collection.driver.FindElement(By.XPath(element));
            hover.MoveToElement(topic).Build().Perform();
        }

        public static void mouseclick(string element, Property_type type)
        {
            Actions popupclick = new Actions(Property_Collection.driver);
            IWebElement accept = Property_Collection.driver.FindElement(By.XPath(element));
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
        /*  public static void click_Checkbox(String element,String value)
          {

              if (value=="YES")
              Property_Collection.driver.FindElement(By.XPath(element)).Click();
              


          }*/

        public static string GetText(string element, Property_type type)
        {
            if (type == Property_type.Id)
                return Property_Collection.driver.FindElement(By.Id(element)).GetAttribute("value");
            if (type == Property_type.XPath)
                return Property_Collection.driver.FindElement(By.XPath(element)).GetAttribute("value");
            else return String.Empty;
        }
        public static string GetTextFromDDL(string element, Property_type type)
        {
            if (type == Property_type.Id)
                return new SelectElement(Property_Collection.driver.FindElement(By.Id(element))).AllSelectedOptions.SingleOrDefault().Text;
            if (type == Property_type.XPath)
                return new SelectElement(Property_Collection.driver.FindElement(By.XPath(element))).AllSelectedOptions.SingleOrDefault().Text;
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

       /* public static void Vehicletype(String element, String value, String costnew,String vtype,String hp)
        {
            if (value == "CAR")
                new SelectElement(Property_Collection.driver.FindElement(By.XPath(element))).SelectByText(value);
            if (value == "PICKUP W/O CAMPER")
                new SelectElement(Property_Collection.driver.FindElement(By.XPath(element))).SelectByText(value);
            if (value == "SUV")
                new SelectElement(Property_Collection.driver.FindElement(By.XPath(element))).SelectByText(value);
            if (value == "VAN")
                new SelectElement(Property_Collection.driver.FindElement(By.XPath(element))).SelectByText(value);

            if (value == "MOTORCYCLE")
            {
                new SelectElement(Property_Collection.driver.FindElement(By.XPath(element))).SelectByText(value);
                System.Threading.Thread.Sleep(1000);
                Console.WriteLine("Motorcycle Selected");
                Property_Collection.driver.FindElement(By.XPath("//input[contains(@id,'txtCostNew_1')]")).Clear();

                Property_Collection.driver.FindElement(By.XPath("//input[contains(@id,'txtCostNew_1')]")).SendKeys(costnew);
                Console.WriteLine("Cost New Entered");
                new SelectElement(Property_Collection.driver.FindElement(By.XPath("//select[contains(@id,'MotorCyleType_1')]"))).SelectByText(vtype);
                Console.WriteLine("Vehicle type selected");
                Property_Collection.driver.FindElement(By.XPath(".//input[contains(@id,'_txtHorsePower_1')]")).Clear();
                Property_Collection.driver.FindElement(By.XPath(".//input[contains(@id,'_txtHorsePower_1')]")).SendKeys(hp);
                Console.WriteLine("Horsepower entered");
                System.Threading.Thread.Sleep(1000);
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
            Property_Collection.driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);
        }

        public static void ScreenShot(string location)
        {
            ITakesScreenshot ss = Property_Collection.driver as ITakesScreenshot;
            Screenshot screenshot = ss.GetScreenshot();
           // DateTime time = DateTime.Now;
          // string dateToday = "_date_" + time.ToString("yyyy-MM-dd") + "_time_" + time.ToString("HH-mm-ss");
            screenshot.SaveAsFile(location, ScreenshotImageFormat.Png);
        }
        public static void CheckTitle(String ExpectedTitle)
        {
            if (Property_Collection.driver.Title.Equals(ExpectedTitle))
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
            if (Property_Collection.driver.PageSource.Contains(ExpectedText))
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
                if (Property_Collection.driver.FindElement(By.XPath(element)).Displayed == true)
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







    



