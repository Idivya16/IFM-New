using NUnit.Framework;
using NUnit.Framework.Interfaces;
using OpenQA.Selenium;
using RelevantCodes.ExtentReports;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace VR_HOME
{
    public class NB_HOME
    {
        String url;
        String path = @"C:\Users\imnay\Documents\Visual Studio 2015\Projects\VR_HOME\VR_HOME\Output\";
        String ReportPath = @"C:\Users\imnay\Documents\Visual Studio 2015\Projects\VR_HOME\VR_HOME\Report\";
        [SetUp]
        public void Initialize()

        {
            Perform.Browser("chrome");
          url = "http://www.ifmig.net/NewPublicSite/NewPublicHome.aspx";
            //url = "http://ifmoldrelease/NewPublicSite/NewPublicHome.aspx";
            ExcelUtil.setExcelFile(@"C:\Users\imnay\Documents\Visual Studio 2015\Projects\VR_HOME\VR_HOME\Excel\VR_HOM_NB01.xlsx");
            Perform.report = new ExtentReports(ReportPath + "Report.html", CultureInfo.GetCultureInfo("es-ES"), true, DisplayOrder.NewestFirst);
            //Perform.report.LoadConfig(ReportPath + "extent-config.xml");
        }

        [Test]
        public void LoginPage()
        {
            Perform.test = Perform.report.StartTest("Login Page");

            //Login Page
            Perform.driver.Navigate().GoToUrl(url);
            Console.WriteLine("Browser Opened");
            //Click on Agents Only Link
            Perform.Click(".//*[@id='Footer_AgentsOnlyLink']");

            //Page 1
            //Enter Username
            Perform.EnterText(".//*[@id='Application_txtUsername']", "DonBrewtonTest");
            //Enter Password
            Perform.EnterText(".//*[@id='Application_txtPassword']", "DonBrewtonTest1");
            Perform.test.Log(LogStatus.Info, "User Credentials Entered");
            Perform.Click(".//*[@id='Application_btnLogin']");
            Perform.test.Log(LogStatus.Info, "Login button Clicked");
            Perform.CheckTitle("Welcome Agents!");

        }
        [Test]
        public void Menu()
        {
            LoginPage();
            Perform.test = Perform.report.StartTest("Menu Page");
            //Page 2
            //Select Velocirater from Agency Info
            Perform.mouseHover(".//*[@id='ulaitem0_0']");
            Perform.waitTillElementToAppear("//*[@id='ulaitem0_0_1']");
            // Perform.Wait();
            Perform.Click("//*[@id='ulaitem0_0_1']");
            Console.WriteLine("VelociRater is Clicked");
            Perform.test.Log(LogStatus.Info, "Velocirater Clicked");
            Perform.CheckTitle("VelociRater");

        }
        [Test]
        public void LOBPage()
        {
            Menu();
            Perform.test = Perform.report.StartTest("LOB Page");
            //Select New Property/Liability Quote
            Perform.Click(".//*[@id='main']/table/tbody/tr/td[2]/div[2]/div[2]/div/input");
            //Select Personal Home
            Perform.Click(".//*[@id='ui-id-2']/a");
            Console.WriteLine("New Personal Home is selected");
            Perform.test.Log(LogStatus.Info, "New Personal Home is Selected");
            Perform.PageContains("Underwriting Questions");
        }
        [Test]
        public void UnderwritingPopUp(int row, string sheetname, string savescreenshot)
        {
            LOBPage();
            Perform.test = Perform.report.StartTest("General Underwriting Page");
            //Select NO to all questions
            Perform.click_on_webElements("//input[contains(@id,'radNo_')]", Perform.driver);
            Perform.Click(".//*[@id='radHOMMultiPolicyNo']");
            Perform.test.Log(LogStatus.Info, "Underwriting questions answered as NO");
            //Home Form Selection
            Perform.Click(".//*[@id='cphMain_ctlUWQuestionsPopup_ddHomFormList']");
            if(ExcelUtil.GetCellData(row,0,sheetname)== "HO-2 - HOMEOWNERS BROAD FORM")
            {
                Perform.Click(".//*[@id='cphMain_ctlUWQuestionsPopup_ddHomFormList']/option[1]");
            }
            if (ExcelUtil.GetCellData(row, 0, sheetname) == "HO-3 - HOMEOWNERS SPECIAL FORM")
            {
                Perform.Click(".//*[@id='cphMain_ctlUWQuestionsPopup_ddHomFormList']/option[2]");
            }
            if (ExcelUtil.GetCellData(row, 0, sheetname) == "HO-3 - with HO-15")
            {
                Perform.Click(".//*[@id='cphMain_ctlUWQuestionsPopup_ddHomFormList']/option[3]");
            }
            if (ExcelUtil.GetCellData(row, 0, sheetname) == "HO-4 - HOMEOWNERS CONTENTS BROAD FORM")
            {
                Perform.Click(".//*[@id='cphMain_ctlUWQuestionsPopup_ddHomFormList']/option[4]");
            }
            if (ExcelUtil.GetCellData(row, 0, sheetname) == "HO-6 - HOMEOWNERS UNIT OWNERS FORM")
            {
                Perform.Click(".//*[@id='cphMain_ctlUWQuestionsPopup_ddHomFormList']/option[5]");
            }
            if (ExcelUtil.GetCellData(row, 0, sheetname) == "ML-2 - MOBILE HOME OWNER OCCUPIED")
            {
                Perform.Click(".//*[@id='cphMain_ctlUWQuestionsPopup_ddHomFormList']/option[6]");
            }
            if (ExcelUtil.GetCellData(row, 0, sheetname) == "ML-4 - MOBILE HOME TENANT OCCUPIED")
            {
                Perform.Click(".//*[@id='cphMain_ctlUWQuestionsPopup_ddHomFormList']/option[7]");
            }
            // Perform.SelectDropDown(".//*[@id='cphMain_ctlUWQuestionsPopup_ddHomFormList']", ExcelUtil.GetCellData(row, 0, sheetname));

            Console.WriteLine("Underwriting questions answered");
            Perform.test.Log(LogStatus.Info, "Home Type is Selected");
            Perform.Click(".//*[@id='cphMain_ctlUWQuestionsPopup_btnSave']");
            Perform.PageContains("*First Name");

        }
        [Test]
        public void PolicyHolderPage(int row, string sheetname, string savescreenshot)
        {
            UnderwritingPopUp(row, "FORM", savescreenshot);
            Perform.test = Perform.report.StartTest("PolicyHolder Page");
            Perform.test.Log(LogStatus.Info, "Enter Policyholder Information");
            //PolicyHolder Page
            //PolicyHolder1 Details
            Perform.EnterText(".//*[@id='cphMain_ctlHomeInput_ctlIsuredList_ctlInsured_txtFirstName']", ExcelUtil.GetCellData(row, 0, sheetname));
            Perform.EnterText(".//*[@id='cphMain_ctlHomeInput_ctlIsuredList_ctlInsured_txtMiddleName']", ExcelUtil.GetCellData(row, 1, sheetname));
            Perform.EnterText(".//*[@id='cphMain_ctlHomeInput_ctlIsuredList_ctlInsured_txtLastName']", ExcelUtil.GetCellData(row, 2, sheetname));
            Perform.SelectDropDown(".//*[@id='cphMain_ctlHomeInput_ctlIsuredList_ctlInsured_ddSuffix']", ExcelUtil.GetCellData(row, 3, sheetname));
            Perform.SelectDropDown(".//*[@id='cphMain_ctlHomeInput_ctlIsuredList_ctlInsured_ddSex']", ExcelUtil.GetCellData(row, 4, sheetname));
            Perform.EnterText(".//*[@id='cphMain_ctlHomeInput_ctlIsuredList_ctlInsured_txtSSN']", ExcelUtil.GetCellData(row, 5, sheetname));
            Perform.EnterText(".//*[@id='cphMain_ctlHomeInput_ctlIsuredList_ctlInsured_txtBirthDate']", ExcelUtil.GetCellData(row, 6, sheetname));
            Perform.EnterText(".//*[@id='cphMain_ctlHomeInput_ctlIsuredList_ctlInsured_txtEmail']", ExcelUtil.GetCellData(row, 7, sheetname));
            Perform.EnterText(".//*[@id='cphMain_ctlHomeInput_ctlIsuredList_ctlInsured_txtPhone']", ExcelUtil.GetCellData(row, 8, sheetname));
            Perform.EnterText(".//*[@id='cphMain_ctlHomeInput_ctlIsuredList_ctlInsured_txtPhoneExt']", ExcelUtil.GetCellData(row, 9, sheetname));
            Perform.SelectDropDown(".//*[@id='cphMain_ctlHomeInput_ctlIsuredList_ctlInsured_ddPhoneType']", ExcelUtil.GetCellData(row, 10, sheetname));
            Perform.EnterText(".//*[@id='cphMain_ctlHomeInput_ctlIsuredList_ctlInsured_txtStreetNum']", ExcelUtil.GetCellData(row, 11, sheetname));
            Perform.EnterText(".//*[@id='cphMain_ctlHomeInput_ctlIsuredList_ctlInsured_txtStreetName']", ExcelUtil.GetCellData(row, 12, sheetname));
            Perform.EnterText(".//*[@id='cphMain_ctlHomeInput_ctlIsuredList_ctlInsured_txtAptNum']", ExcelUtil.GetCellData(row, 13, sheetname));
            Perform.EnterText(".//*[@id='cphMain_ctlHomeInput_ctlIsuredList_ctlInsured_txtPOBox']", ExcelUtil.GetCellData(row, 14, sheetname));
            Perform.EnterText(".//*[@id='cphMain_ctlHomeInput_ctlIsuredList_ctlInsured_txtZipCode']", ExcelUtil.GetCellData(row, 15, sheetname));
            System.Threading.Thread.Sleep(500);
            Perform.EnterText(".//*[@id='cphMain_ctlHomeInput_ctlIsuredList_ctlInsured_txtCityName']", ExcelUtil.GetCellData(row, 16, sheetname));
            //Perform.Wait();

            // Perform.EnterText(".//*[@id='cphMain_ctl_Master_Edit_ctlIsuredList_ctlInsured_txtCityName']", ExcelUtil.GetCellData(row, 16, Sheetname), Property_type.XPath);
            // Perform.SelectDropDown(".//*[@id='cphMain_ctlHomeInput_ctlIsuredList_ctlInsured_ddStateAbbrev']", ExcelUtil.GetCellData(row, 17, sheetname));
            Perform.driver.FindElement(By.XPath(".//*[@id='cphMain_ctlHomeInput_ctlIsuredList_ctlInsured_txtGaragedCounty']")).Clear();
            Perform.EnterText(".//*[@id='cphMain_ctlHomeInput_ctlIsuredList_ctlInsured_txtGaragedCounty']", ExcelUtil.GetCellData(row, 18, sheetname));
            Console.WriteLine("Policyholder1 details are entered");

            //PolicyHolder2 Details
            if (ExcelUtil.GetCellData(row, 19, sheetname) != "")
            {
                Perform.Click(".//*[@id='cphMain_ctlHomeInput_ctlIsuredList_ctlInsured1_lblInsuredTitle']");
                Perform.waitTillElementToAppear(".//*[@id='cphMain_ctlHomeInput_ctlIsuredList_ctlInsured1_txtFirstName']");
                Perform.EnterText(".//*[@id='cphMain_ctlHomeInput_ctlIsuredList_ctlInsured1_txtFirstName']", ExcelUtil.GetCellData(row, 19, sheetname));
                Perform.EnterText(".//*[@id='cphMain_ctlHomeInput_ctlIsuredList_ctlInsured1_txtMiddleName']", ExcelUtil.GetCellData(row, 20, sheetname));
                Perform.EnterText(".//*[@id='cphMain_ctlHomeInput_ctlIsuredList_ctlInsured1_txtLastName']", ExcelUtil.GetCellData(row, 21, sheetname));
                Perform.SelectDropDown(".//*[@id='cphMain_ctlHomeInput_ctlIsuredList_ctlInsured1_ddSuffix']", ExcelUtil.GetCellData(row, 22, sheetname));
                Perform.SelectDropDown(".//*[@id='cphMain_ctlHomeInput_ctlIsuredList_ctlInsured1_ddSex']", ExcelUtil.GetCellData(row, 23, sheetname));
                Perform.EnterText(".//*[@id='cphMain_ctlHomeInput_ctlIsuredList_ctlInsured1_txtSSN']", ExcelUtil.GetCellData(row, 24, sheetname));
                Perform.EnterText(".//*[@id='cphMain_ctlHomeInput_ctlIsuredList_ctlInsured1_txtBirthDate']", ExcelUtil.GetCellData(row, 25, sheetname));
                Perform.EnterText(".//*[@id='cphMain_ctlHomeInput_ctlIsuredList_ctlInsured1_txtEmail']", ExcelUtil.GetCellData(row, 26, sheetname));
                Perform.EnterText(".//*[@id='cphMain_ctlHomeInput_ctlIsuredList_ctlInsured1_txtPhone']", ExcelUtil.GetCellData(row, 27, sheetname));
                Perform.EnterText(".//*[@id='cphMain_ctlHomeInput_ctlIsuredList_ctlInsured1_txtPhoneExt']", ExcelUtil.GetCellData(row, 28, sheetname));
                Perform.SelectDropDown(".//*[@id='cphMain_ctlHomeInput_ctlIsuredList_ctlInsured1_ddPhoneType']", ExcelUtil.GetCellData(row, 29, sheetname));
                Console.WriteLine("Policy holder 2 details are entered");
            }
            //Take Screenshot
            Perform.ScreenShot(savescreenshot + "PolicyHolderPage.png");
            Perform.test.Log(LogStatus.Info, "Policyholder details are entered");
            Perform.test.Log(LogStatus.Info, "Proceed to Property Page");
            Perform.Click(".//*[@id='cphMain_ctlHomeInput_ctlIsuredList_btnSaveAndGotoDrivers']");
            Perform.IsElementPresent(".//*[@id='cphMain_ctlHomeInput_ctlProperty_HOM_btnSaveGotoNextSection']");
        }
        [Test]
        public void PropertyPage(int row,string sheetname,string savescreenshot)
        {
            PolicyHolderPage(row, "POLICYHOLDER", savescreenshot);
            Perform.test = Perform.report.StartTest("Property Page");
            Perform.test.Log(LogStatus.Info, "Enter Property Information");
            //Address Details

            Perform.Click(".//*[@id='ui-id-23']");
            if(ExcelUtil.GetCellData(row,0,"FORM")== "ML-2 - MOBILE HOME OWNER OCCUPIED")
            {
                Perform.Click(".//*[@id='ui-id-23']");
            }
           // Perform.Click(".//*[@id='cphMain_ctlHomeInput_ctlProperty_HOM_ctlProperty_Address_lblAccordHeader']");
            Perform.waitTillElementToAppear(".//*[@id='cphMain_ctlHomeInput_ctlProperty_HOM_ctlProperty_Address_txtStreetNum']");
            Perform.EnterText(".//*[@id='cphMain_ctlHomeInput_ctlProperty_HOM_ctlProperty_Address_txtStreetNum']", ExcelUtil.GetCellData(row,0, sheetname));
            Perform.EnterText(".//*[@id='cphMain_ctlHomeInput_ctlProperty_HOM_ctlProperty_Address_txtStreetName']", ExcelUtil.GetCellData(row,1, sheetname));
            Perform.EnterText(".//*[@id='cphMain_ctlHomeInput_ctlProperty_HOM_ctlProperty_Address_txtAptNum']", ExcelUtil.GetCellData(row,2, sheetname));
            Perform.EnterText(".//*[@id='cphMain_ctlHomeInput_ctlProperty_HOM_ctlProperty_Address_txtZipCode']", ExcelUtil.GetCellData(row,3, sheetname));
           // Perform.EnterText(".//*[@id='cphMain_ctlHomeInput_ctlProperty_HOM_ctlProperty_Address_txtCityName']", ExcelUtil.GetCellData(row,4, sheetname));
           // Perform.SelectDropDown(".//*[@id='cphMain_ctlHomeInput_ctlProperty_HOM_ctlProperty_Address_ddStateAbbrev']", ExcelUtil.GetCellData(row,5, sheetname));
            Perform.driver.FindElement(By.XPath(".//*[@id='cphMain_ctlHomeInput_ctlProperty_HOM_ctlProperty_Address_txtGaragedCounty']")).Clear();
            Perform.EnterText(".//*[@id='cphMain_ctlHomeInput_ctlProperty_HOM_ctlProperty_Address_txtGaragedCounty']", ExcelUtil.GetCellData(row,6, sheetname));
            Perform.test.Log(LogStatus.Info, "Address Entered");
            //Residence Details
            Perform.Wait();
            Perform.EnterText(".//*[@id='txtYearBuilt0']", ExcelUtil.GetCellData(row,7, sheetname));
            Perform.EnterText(".//*[@id='cphMain_ctlHomeInput_ctlProperty_HOM_ctlResidence_txtSqrFeet']", ExcelUtil.GetCellData(row,8, sheetname));
            Perform.SelectDropDown(".//*[@id='cphMain_ctlHomeInput_ctlProperty_HOM_ctlResidence_ddlNumberOfFamilies']", ExcelUtil.GetCellData(row,9, sheetname));
            if (Perform.driver.FindElement(By.XPath(".//*[@id='cphMain_ctlHomeInput_ctlProperty_HOM_ctlResidence_ddlStructure']")).Enabled && ExcelUtil.GetCellData(row, 10, sheetname)!="")
            {
                Perform.SelectDropDown(".//*[@id='cphMain_ctlHomeInput_ctlProperty_HOM_ctlResidence_ddlStructure']", ExcelUtil.GetCellData(row, 10, sheetname));
            }
            if (Perform.driver.FindElement(By.XPath(".//*[@id='cphMain_ctlHomeInput_ctlProperty_HOM_ctlResidence_ddlOccupancy']")).Enabled && ExcelUtil.GetCellData(row, 11, sheetname) != "")
            {
                Perform.SelectDropDown(".//*[@id='cphMain_ctlHomeInput_ctlProperty_HOM_ctlResidence_ddlOccupancy']", ExcelUtil.GetCellData(row, 11, sheetname));
            }
            if (Perform.driver.FindElement(By.XPath(".//*[@id='cphMain_ctlHomeInput_ctlProperty_HOM_ctlResidence_ddlConstruction']")).Enabled && ExcelUtil.GetCellData(row, 12, sheetname) != "")
            {
                Perform.SelectDropDown(".//*[@id='cphMain_ctlHomeInput_ctlProperty_HOM_ctlResidence_ddlConstruction']", ExcelUtil.GetCellData(row, 12, sheetname));
            }
            //if (Perform.driver.FindElement(By.XPath(".//*[@id='cphMain_ctlHomeInput_ctlProperty_HOM_ctlResidence_ddStyle']")).Displayed)
           // {
                if (ExcelUtil.GetCellData(row, 13, sheetname) != "")
                {
                    Perform.SelectDropDown(".//*[@id='cphMain_ctlHomeInput_ctlProperty_HOM_ctlResidence_ddStyle']", ExcelUtil.GetCellData(row, 13, sheetname));
                }
            //  }
            Perform.test.Log(LogStatus.Info, "Residence Details Entered");
            if (ExcelUtil.GetCellData(row,14,sheetname)!="")
            {
                Perform.SelectDropDown(".//*[@id='cphMain_ctlHomeInput_ctlProperty_HOM_ctlMobileHome_ddlTieDown']", ExcelUtil.GetCellData(row, 14, sheetname));
                Perform.SelectDropDown(".//*[@id='cphMain_ctlHomeInput_ctlProperty_HOM_ctlMobileHome_ddlSkirting']", ExcelUtil.GetCellData(row, 15, sheetname));
                Perform.SelectDropDown(".//*[@id='cphMain_ctlHomeInput_ctlProperty_HOM_ctlMobileHome_ddlFoundation']", ExcelUtil.GetCellData(row, 16, sheetname));
                Perform.EnterText(".//*[@id='cphMain_ctlHomeInput_ctlProperty_HOM_ctlMobileHome_txtMake']", ExcelUtil.GetCellData(row, 17, sheetname));
                Perform.EnterText(".//*[@id='cphMain_ctlHomeInput_ctlProperty_HOM_ctlMobileHome_txtModel']", ExcelUtil.GetCellData(row, 18, sheetname));
                Perform.EnterText(".//*[@id='cphMain_ctlHomeInput_ctlProperty_HOM_ctlMobileHome_txtVin']", ExcelUtil.GetCellData(row, 19, sheetname));
            }
            //Protection Class Details
            if (ExcelUtil.GetCellData(row, 20, sheetname) != "")
            {
                Perform.SelectDropDown(".//*[@id='cphMain_ctlHomeInput_ctlProperty_HOM_ctlProtectionClass_HOM_ddlFeetToHydrantC']", ExcelUtil.GetCellData(row, 20, sheetname));
            }
            Perform.driver.FindElement(By.XPath(".//*[@id='cphMain_ctlHomeInput_ctlProperty_HOM_ctlProtectionClass_HOM_txtMilesToFireDepartmentC']")).SendKeys(ExcelUtil.GetCellData(row, 21, sheetname));
            // Perform.EnterText(".//*[@id='cphMain_ctlHomeInput_ctlProperty_HOM_ctlProtectionClass_HOM_txtMilesToFireDepartmentC']", ExcelUtil.GetCellData(row,21, sheetname));
            Perform.test.Log(LogStatus.Info, "Protection Class Details Entered");
            //Additional Questions
            if (ExcelUtil.GetCellData(row,22,sheetname)=="YES" && ExcelUtil.GetCellData(row, 22, sheetname) != "")
            {
                System.Threading.Thread.Sleep(500);
                Perform.waitTillElementToAppear(".//*[@id='cphMain_ctlHomeInput_ctlProperty_HOM_ctlPropertyAdditionalQuestions_chkHasAutoPolicy']");
                Perform.Click(".//*[@id='cphMain_ctlHomeInput_ctlProperty_HOM_ctlPropertyAdditionalQuestions_chkHasAutoPolicy']");
            }
            if (ExcelUtil.GetCellData(row,23, sheetname) == "YES" && ExcelUtil.GetCellData(row, 23, sheetname) != "")
            {
                System.Threading.Thread.Sleep(200);
                Perform.Click(".//*[@id='cphMain_ctlHomeInput_ctlProperty_HOM_ctlPropertyAdditionalQuestions_chkFirstWrittenDate']");
                Perform.EnterText(".//*[@id='cphMain_ctlHomeInput_ctlProperty_HOM_ctlPropertyAdditionalQuestions_txtFirstWrittenDate']", ExcelUtil.GetCellData(row, 24, sheetname));
            }
            if (ExcelUtil.GetCellData(row,25, sheetname) == "YES" && Perform.driver.FindElement(By.XPath(".//*[@id='cphMain_ctlHomeInput_ctlProperty_HOM_ctlPropertyAdditionalQuestions_chkAnyChildren']")).Displayed)
            {
                Perform.Click(".//*[@id='cphMain_ctlHomeInput_ctlProperty_HOM_ctlPropertyAdditionalQuestions_chkAnyChildren']");
            }
            if (ExcelUtil.GetCellData(row,26, sheetname) == "YES" && ExcelUtil.GetCellData(row, 26, sheetname) != "")
            {
                if (Perform.driver.FindElement(By.XPath(".//*[@id='cphMain_ctlHomeInput_ctlProperty_HOM_ctlPropertyAdditionalQuestions_chkSmokeAlarms']")).Displayed)
                {
                    Perform.Click(".//*[@id='cphMain_ctlHomeInput_ctlProperty_HOM_ctlPropertyAdditionalQuestions_chkSmokeAlarms']");
                }
                else
                {
                    Perform.driver.FindElement(By.XPath(".//*[@id='cphMain_ctlHomeInput_ctlProperty_HOM_ctlPropertyAdditionalQuestions_chkSmokeAlarms']")).SendKeys(Keys.PageDown);
                    Perform.Click(".//*[@id='cphMain_ctlHomeInput_ctlProperty_HOM_ctlPropertyAdditionalQuestions_chkSmokeAlarms']");
                }
                    
            }
            Perform.SelectDropDown(".//*[@id='cphMain_ctlHomeInput_ctlProperty_HOM_ctlPropertyAdditionalQuestions_ddBurglarAlarmType']", ExcelUtil.GetCellData(row,27, sheetname));
            Perform.SelectDropDown(".//*[@id='cphMain_ctlHomeInput_ctlProperty_HOM_ctlPropertyAdditionalQuestions_ddFireAlarmType']", ExcelUtil.GetCellData(row,28, sheetname));
            Perform.SelectDropDown(".//*[@id='cphMain_ctlHomeInput_ctlProperty_HOM_ctlPropertyAdditionalQuestions_ddSprinklerType']", ExcelUtil.GetCellData(row,29, sheetname));
            if (ExcelUtil.GetCellData(row,30, sheetname) == "YES" && Perform.driver.FindElement(By.XPath(".//*[@id='cphMain_ctlHomeInput_ctlProperty_HOM_ctlPropertyAdditionalQuestions_chkTrampoline']")).Displayed)
            {
                Perform.Click(".//*[@id='cphMain_ctlHomeInput_ctlProperty_HOM_ctlPropertyAdditionalQuestions_chkTrampoline']");
            }
            if (ExcelUtil.GetCellData(row,31, sheetname) == "YES" && Perform.driver.FindElement(By.XPath(".//*[@id='cphMain_ctlHomeInput_ctlProperty_HOM_ctlPropertyAdditionalQuestions_chkSwimmingPool']")).Displayed)
            {
                Perform.Click(".//*[@id='cphMain_ctlHomeInput_ctlProperty_HOM_ctlPropertyAdditionalQuestions_chkSwimmingPool']");
            }
            if (ExcelUtil.GetCellData(row,32, sheetname) == "YES" && Perform.driver.FindElement(By.XPath(".//*[@id='cphMain_ctlHomeInput_ctlProperty_HOM_ctlPropertyAdditionalQuestions_chkWoodStove']")).Displayed)
            {
                Perform.Click(".//*[@id='cphMain_ctlHomeInput_ctlProperty_HOM_ctlPropertyAdditionalQuestions_chkWoodStove']");

            }
            Perform.test.Log(LogStatus.Info, "Additonal Questions Answered");
            Perform.ScreenShot(savescreenshot + "Property.png");
            Console.WriteLine("Property details are entered");
            Perform.test.Log(LogStatus.Info, "Property Details Entered");
            Perform.Click(".//*[@id='cphMain_ctlHomeInput_ctlProperty_HOM_btnSubmit']");
            Perform.Click(".//*[@id='cphMain_ctlHomeInput_ctlProperty_HOM_btnSaveGotoNextSection']");
            Perform.IsElementPresent(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_btnE2Value']");
        }
        [Test]
        public void CoveragePage(int row,string sheetname,string savescreenshot)
        {
            PropertyPage(row, "PROPERTY", savescreenshot);
            Perform.test = Perform.report.StartTest("Coverage Page");
            Perform.test.Log(LogStatus.Info, "Enter Coverage Information");
            PageUtility.Coveragedetails(row, sheetname);
            Perform.test.Log(LogStatus.Info, "Coverage Details Entered");
            Perform.ScreenShot(savescreenshot + "Coverage.png");
            Perform.Click(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlIMRVWatercraft_ctlInlandMarine_btnRateIM']");
            Perform.IsElementPresent(".//*[@id='cphMain_ctlHomeInput_ctlQuoteSummary_HOM_ctlQuoteSummaryActions_btnContinueToApp']");
        }
        [Test]
        public void QuoteSummaryPage(int row,string sheetname,string savescreenshot)
        {
            CoveragePage(row, "COVERAGES", savescreenshot);
            Perform.test = Perform.report.StartTest("Quote Summary Page");
           
            //Click Continue to Application
            Perform.ScreenShot(savescreenshot + "QuoteSummary.png");
            Perform.Click(".//*[@id='cphMain_ctlHomeInput_ctlQuoteSummary_HOM_ctlQuoteSummaryActions_btnContinueToApp']");
            Perform.test.Log(LogStatus.Info, "Continue to Application is Clicked");
            Perform.IsElementPresent(".//*[@id='cphMain_ctl_Master_HOM_APP_ctlUWQuestions_btnGoToApp']");
        }
        [Test]
        public void UnderwritingPage(int row,string sheetname,string savescreenshot)
        {
            QuoteSummaryPage(row, "", savescreenshot);
            Perform.test = Perform.report.StartTest("Underwriting Questions Page");
            Perform.Click("//input[contains(@id,'rbNo_0')]");
            Perform.Click("//input[contains(@id,'rbNo_1')]");
            Perform.Click("//input[contains(@id,'rbNo_2')]");
            Perform.Click("//input[contains(@id,'rbNo_3')]");
            Perform.Click("//input[contains(@id,'rbNo_4')]");
            Perform.Click("//input[contains(@id,'rbNo_5')]");

            Perform.Click("//input[contains(@id,'rbNo_10')]");
            Perform.Click("//input[contains(@id,'rbNo_11')]");
            Perform.Click("//input[contains(@id,'rbNo_14')]");
            Perform.Click("//input[contains(@id,'rbNo_15')]");
            Perform.Click("//input[contains(@id,'rbNo_16')]");
            Perform.Click("//input[contains(@id,'rbNo_17')]");
            Perform.Click("//input[contains(@id,'rbNo_18')]");
            Perform.Click("//input[contains(@id,'rbNo_19')]");
            Perform.Click("//input[contains(@id,'rbNo_20')]");
            Perform.Click("//input[contains(@id,'rbNo_21')]");
            Perform.Click("//input[contains(@id,'rbNo_22')]");
            if (ExcelUtil.GetCellData(row, 0, "FORM") == "HO-4 - HOMEOWNERS CONTENTS BROAD FORM" || ExcelUtil.GetCellData(row, 0, "FORM") == "ML-4 - MOBILE HOME TENANT OCCUPIED" || ExcelUtil.GetCellData(row, 0, "FORM") == "HO-6 - HOMEOWNERS UNIT OWNERS FORM")
            {
                Perform.Click("//input[contains(@id,'rbNo_23')]");
                Perform.Click("//input[contains(@id,'rbNo_24')]");
                Perform.Click("//input[contains(@id,'rbNo_25')]");
               
            }
                Console.WriteLine("Underwriting Questions answered");
            Perform.test.Log(LogStatus.Info, "Underwriting Questions Answered");
            Perform.ScreenShot(savescreenshot + "Underwriting.png");
            Perform.Click(".//*[@id='cphMain_ctl_Master_HOM_APP_ctlUWQuestions_btnGoToApp']");
            Perform.IsElementPresent(".//*[@id='ui-id-23']");
        }
        [Test]
        public void ApplicationPage(int row,string sheetname,string savescreenshot)
        {
            UnderwritingPage(row, "", savescreenshot);
            Perform.test = Perform.report.StartTest("Application Page");
            PageUtility.Application(row, sheetname);
            Perform.test.Log(LogStatus.Info, "Property Updates Given");
            //Select Billing Info
            Perform.SelectDropDown(".//*[@id='cphMain_ctl_Master_HOM_APP_ctl_HOM_App_Section_ctl_Billing_Info_PPA_ddMethod']", ExcelUtil.GetCellData(row, 0, sheetname));
            Perform.SelectDropDown(".//*[@id='cphMain_ctl_Master_HOM_APP_ctl_HOM_App_Section_ctl_Billing_Info_PPA_ddPayPlan']", ExcelUtil.GetCellData(row, 1, sheetname));
            Perform.SelectDropDown(".//*[@id='cphMain_ctl_Master_HOM_APP_ctl_HOM_App_Section_ctl_Billing_Info_PPA_ddBillTo']", ExcelUtil.GetCellData(row, 2, sheetname));
            Perform.test.Log(LogStatus.Info, "Billing Info is Selected");
            Perform.Click(".//*[@id='btnShowEffectiveDate']");
            //Choose Effective Date
            Perform.SelectDropDown(".//*[@id='ui-datepicker-div']/div[1]/div/select[1]", ExcelUtil.GetCellData(row, 3, sheetname));
            Console.WriteLine("Month selected");
            Perform.SelectDropDown(".//*[@id='ui-datepicker-div']/div[1]/div/select[2]", ExcelUtil.GetCellData(row, 5, sheetname));
            Console.WriteLine("Year Selected");
           string date = ExcelUtil.GetCellData(row,4, sheetname);
            Console.WriteLine(date);
            Perform.driver.FindElement(By.LinkText(date)).Click();
            Console.WriteLine("Day selected");
            Perform.driver.FindElement(By.Id("btnEffectiveDateDone")).SendKeys(Keys.Enter);
            Perform.Wait();
            Console.WriteLine("Effective Date Entered");
            Perform.test.Log(LogStatus.Info, "Effective Date Entered");

        }
        

    
        [Test]
        public void FinalizePage()
        {

            int sheetrownum = ExcelUtil.getRowCount("TestCase");
            try
            {
                for (int i =25; i < sheetrownum; i++)
                {
                    System.IO.Directory.CreateDirectory(path + ExcelUtil.GetCellData(i, 0, "TestCase"));
                    string savescreenshot = path + ExcelUtil.GetCellData(i, 0, "TestCase") + "\\";
                    ApplicationPage(i, "Billing_Information", savescreenshot);
                    Perform.test = Perform.report.StartTest("Home Policy Issue");
                    Perform.Click(".//*[@id='cphMain_ctl_Master_HOM_APP_ctlQuoteSummary_HOM_ctlQuoteSummaryActions_btnContinueToApp']");
                    Perform.Click(".//*[@id='cphMain_ctl_Master_HOM_APP_ctlQuoteSummary_HOM_ctlQuoteSummaryActions_lnkFinalize']");
                    Console.WriteLine("Application Finalized");
                    Perform.test.Log(LogStatus.Info, "Application Finalized");
                    if (ExcelUtil.GetCellData(i,0,"RV_WATERCRAFT")!="")
                    {
                        Perform.driver.SwitchTo().Alert().Accept();
                    }
                    Perform.ScreenShot(savescreenshot + "Policy" + (i - 2) + "_" + DateTime.Now.ToString("yyyyMMddhhmmss") + ".png");
                    Perform.test.Log(LogStatus.Info, "Policy Issued");
                    Perform.Click(".//*[@id='CrumbsLogoutLink']");
                    Perform.test.Log(LogStatus.Info, "Logged Out");
                    Console.WriteLine("Policy" + (i - 2) + " issued");

                }
            }catch(Exception e)
            {
                Console.WriteLine("Error:" + e);
                Assert.True(false);
            }
            }
        [TearDown]
        public void CleanUp()
        {
            if (TestContext.CurrentContext.Result.Outcome != ResultState.Success)

                Perform.ScreenShot(path + "Failure.png");
            Perform.driver.Close();
            Perform.report.EndTest(Perform.test);
            Perform.report.Flush();
        }
    }
}
