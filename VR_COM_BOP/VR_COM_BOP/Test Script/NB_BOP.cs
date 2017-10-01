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

namespace VR_COM_BOP
{
    public class NB_BOP
    {
        String url;
        String path = @"C:\Users\imnay\Documents\Visual Studio 2015\Projects\VR_COM_BOP\VR_COM_BOP\Output\";
        
        String ReportPath = @"C:\Users\imnay\Documents\Visual Studio 2015\Projects\VR_COM_BOP\VR_COM_BOP\Report\";

        [SetUp]
        public void Initialize()

        {
            Perform.Browser("firefox");
         //url = "http://ifmnewrelease/NewPublicSite/NewPublicHome.aspx";
           // url = "http://ifmoldrelease/NewPublicSite/NewPublicHome.aspx";
        url = "http://www.ifmig.net/NewPublicSite/NewPublicHome.aspx";
            ExcelUtil.setExcelFile(@"C:\Users\imnay\Documents\Visual Studio 2015\Projects\VR_COM_BOP\VR_COM_BOP\Excel\VR_NB_BOP.xlsx");

            Perform.report = new ExtentReports(ReportPath+"Report.html",CultureInfo.GetCultureInfo("es-ES"),true,DisplayOrder.NewestFirst);
            Perform.report.LoadConfig(ReportPath + "extent-config.xml");
           
            
        }

        [Test]
        public void LoginPage()
        {
            Perform.test = Perform.report.StartTest("Login Page");
            
            //Login Page
            Perform.driver.Navigate().GoToUrl(url);
            Console.WriteLine("Browser Opened");
            Perform.test.Log(LogStatus.Info,"Browser Opened");
            //Click on Agents Only Link
            Perform.driver.FindElement(By.XPath(".//*[@id='Footer_AgentsOnlyLink']")).SendKeys(Keys.PageDown);
            Perform.Click(".//*[@id='Footer_AgentsOnlyLink']", Property_type.XPath);

            //Page 1
            //Enter Username
            Perform.EnterText(".//*[@id='Application_txtUsername']", "DonBrewtonTest", Property_type.XPath);
            Perform.test.Log(LogStatus.Info, "Username Entered");
            //Enter Password
            Perform.EnterText(".//*[@id='Application_txtPassword']", "DonBrewtonTest1", Property_type.XPath);
            Perform.test.Log(LogStatus.Info, "Password Entered");
            Perform.Click(".//*[@id='Application_btnLogin']", Property_type.XPath);
            Perform.test.Log(LogStatus.Info, "Login button Clicked");
            Perform.CheckTitle("Welcome Agents!");


        }
        [Test]
        public void Menu()
        {
            LoginPage();

            //Page 2
            //Select Velocirater from Agency Info
            Perform.test = Perform.report.StartTest("Main Page");
            //Perform.waitTillElementToAppear(".//*[@id='ulaitem0_0']");
            // Perform.mouseHover(".//*[@id='ulaitem0_0']", Property_type.XPath);
            Perform.driver.Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(10);
            Perform.Click(".//*[@id='ulaitem0_0']", Property_type.XPath);
            System.Threading.Thread.Sleep(2000);
            Perform.waitTillElementToAppear("//*[@id='ulaitem0_0_1']");
            Perform.Click("//*[@id='ulaitem0_0_1']", Property_type.XPath);
            Console.WriteLine("VelociRater is Clicked");
            Perform.test.Log(LogStatus.Info, "Velocirater Clicked");
            Perform.CheckTitle("VelociRater");

        }
        [Test]
        public void LOBPage()
        {
            Menu();
            
            //Select New Commercial Quote
            Perform.test = Perform.report.StartTest("LOB Page");
            Perform.driver.Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(20);
            Perform.waitTillElementToAppear(".//*[@id='main']/table/tbody/tr/td[2]/div[3]/div[2]/div/input");
            Perform.Click(".//*[@id='main']/table/tbody/tr/td[2]/div[3]/div[2]/div/input", Property_type.XPath);
            Perform.test.Log(LogStatus.Info, "New Commercial Quote Clicked");
            //Select Commercial BOP
            Perform.Click(".//*[@id='ui-id-6']/a", Property_type.XPath);
            Perform.test.Log(LogStatus.Info, "Commerical BOP Clicked");
            Perform.PageContains("Underwriting Questions");

        }
        [Test]
        public void UnderwritingPopUp(int row, string sheetname, string savescreenshot)
        {
            LOBPage();
            Perform.driver.Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(10);
            Perform.test = Perform.report.StartTest("General Underwriting Page");
            //Select NO to all questions
            Perform.click_on_webElements("//input[contains(@id,'radNo_')]", Perform.driver);

            Console.WriteLine("Underwriting questions answered");
            Perform.test.Log(LogStatus.Info, "Underwriting questions answered as NO");

            Perform.Click(".//*[@id='cphMain_ctlUWQuestionsPopup_btnSave']", Property_type.XPath);
            Perform.test.Log(LogStatus.Info, "Continue to Quote Clicked");
            Perform.PageContains("Risk Grade Lookup");

        }
        [Test]
        public void RiskGrade(int row, string sheetname, string savescreenshot)
        {
            UnderwritingPopUp(row, "", savescreenshot);
            Perform.driver.Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(10);
            Perform.test = Perform.report.StartTest("Risk Grade Page");
            //Select FilterBy field
            Perform.SelectDropDown(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctlRiskGradeSearch_ddlRiskGradeFilterBy']", ExcelUtil.GetCellData(row, 0, sheetname), Property_type.XPath);
            Perform.test.Log(LogStatus.Info, "Filter By Info Selected");
            //Enter code
            Perform.EnterText(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctlRiskGradeSearch_txtRiskGradeFilterValue']", ExcelUtil.GetCellData(row, 1, sheetname), Property_type.XPath);
            Perform.test.Log(LogStatus.Info, "Filter value Entered");
            Perform.Click(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctlRiskGradeSearch_btnRisksearch']", Property_type.XPath);
            Perform.test.Log(LogStatus.Info, "Find based on the filter values");
            //Select risk grade
            Perform.Click(".//*[@id='DataTables_Table_0']/tbody/tr/td[1]/input", Property_type.XPath);
            Perform.test.Log(LogStatus.Info, "Select the Risk Grade");
            Perform.ScreenShot(savescreenshot + "RiskGradeLookup.png");
            Perform.IsElementPresent(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctlIsuredList_btnSaveAndGotoDrivers']");
        }
        [Test]
        public void Policyholder(int row, string sheetname, string savescreenshot)
        {
            RiskGrade(row, "Risk Look Up", savescreenshot);
            Perform.driver.Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(10);
            Perform.test = Perform.report.StartTest("PolicyHolder Page");
            Perform.EnterText(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctlIsuredList_ctlInsured_txtBusinessName']", ExcelUtil.GetCellData(row, 0, sheetname), Property_type.XPath);
            Perform.EnterText(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctlIsuredList_ctlInsured_txtDBA']", ExcelUtil.GetCellData(row, 1, sheetname), Property_type.XPath);
            Perform.SelectDropDown(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctlIsuredList_ctlInsured_ddBusinessType']", ExcelUtil.GetCellData(row, 2, sheetname), Property_type.XPath);
            Perform.SelectDropDown(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctlIsuredList_ctlInsured_ddTaxIDType']", ExcelUtil.GetCellData(row, 3, sheetname), Property_type.XPath);
            //Check if FEID/SSN
            if (ExcelUtil.GetCellData(row, 3, sheetname) == "FEIN")
                Perform.EnterText(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctlIsuredList_ctlInsured_txtFEIN']", ExcelUtil.GetCellData(row, 4, sheetname), Property_type.XPath);
            if (ExcelUtil.GetCellData(row, 3, sheetname) == "SSN")
                Perform.EnterText(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctlIsuredList_ctlInsured_txtSSNBusiness']", ExcelUtil.GetCellData(row, 5, sheetname), Property_type.XPath);
            Perform.EnterText(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctlIsuredList_ctlInsured_txtBusinessStarted']", ExcelUtil.GetCellData(row, 6, sheetname), Property_type.XPath);
            Perform.EnterText(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctlIsuredList_ctlInsured_txtDescriptionOfOperations']", ExcelUtil.GetCellData(row, 7, sheetname), Property_type.XPath);
            Perform.EnterText(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctlIsuredList_ctlInsured_txtEmail']", ExcelUtil.GetCellData(row, 8, sheetname), Property_type.XPath);
            Perform.EnterText(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctlIsuredList_ctlInsured_txtPhone']", ExcelUtil.GetCellData(row, 9, sheetname), Property_type.XPath);
            Perform.EnterText(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctlIsuredList_ctlInsured_txtPhoneExt']", ExcelUtil.GetCellData(row, 10, sheetname), Property_type.XPath);
            Perform.SelectDropDown(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctlIsuredList_ctlInsured_ddPhoneType']", ExcelUtil.GetCellData(row, 11, sheetname), Property_type.XPath);
            Perform.EnterText(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctlIsuredList_ctlInsured_txtStreetNum']", ExcelUtil.GetCellData(row, 12, sheetname), Property_type.XPath);
            Perform.EnterText(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctlIsuredList_ctlInsured_txtStreetName']", ExcelUtil.GetCellData(row, 13, sheetname), Property_type.XPath);
            Perform.EnterText(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctlIsuredList_ctlInsured_txtAptNum']", ExcelUtil.GetCellData(row, 14, sheetname), Property_type.XPath);
            Perform.EnterText(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctlIsuredList_ctlInsured_txtPOBox']", ExcelUtil.GetCellData(row, 15, sheetname), Property_type.XPath);

            Perform.EnterText(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctlIsuredList_ctlInsured_txtZipCode']", ExcelUtil.GetCellData(row, 16, sheetname), Property_type.XPath);

            //Perform.Wait();
            System.Threading.Thread.Sleep(500);
            Perform.EnterText(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctlIsuredList_ctlInsured_txtCityName']", ExcelUtil.GetCellData(row, 17, sheetname), Property_type.XPath);
            Perform.SelectDropDown(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctlIsuredList_ctlInsured_ddStateAbbrev']", ExcelUtil.GetCellData(row, 18, sheetname), Property_type.XPath);
            Perform.EnterText(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctlIsuredList_ctlInsured_txtGaragedCounty']", ExcelUtil.GetCellData(row, 19, sheetname), Property_type.XPath);
            Console.WriteLine("Policyholder details are entered");
            Perform.test.Log(LogStatus.Info, "Policyholder details are entered");
            Perform.Click(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctlIsuredList_btnSaveAndGotoDrivers']", Property_type.XPath);
            Perform.test.Log(LogStatus.Info, "Proceed to Policy Coverage");
            Perform.ScreenShot(savescreenshot + "PolicyholderDetails.png");
            Perform.IsElementPresent(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_PolicyLevelCoverages_ctl_BOP_Coverages_btnLocations']");
        }
        [Test]
        public void PolicyCoverage(int row, string sheetname, string savescreenshot)
        {
            Policyholder(row, "Policyholder", savescreenshot);
            Perform.driver.Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(10);
            Perform.test = Perform.report.StartTest("Policy Coverage Page");
            Perform.test.Log(LogStatus.Info, "Enter Policy coverage Information");
            //General Information
            Perform.SelectDropDown(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_PolicyLevelCoverages_ctl_BOP_GeneralInformation_ddlOccurrenceLiabilityLimit']", ExcelUtil.GetCellData(row, 0, sheetname), Property_type.XPath);

            Perform.SelectDropDown(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_PolicyLevelCoverages_ctl_BOP_GeneralInformation_ddlTenantsFireLiability']", ExcelUtil.GetCellData(row, 1, sheetname), Property_type.XPath);

            Perform.SelectDropDown(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_PolicyLevelCoverages_ctl_BOP_GeneralInformation_ddlPropertyDamageLiabilityDeductible']", ExcelUtil.GetCellData(row, 2, sheetname), Property_type.XPath);

            Perform.SelectDropDown(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_PolicyLevelCoverages_ctl_BOP_GeneralInformation_ddlPropDmgLiabLimitPerClaimOrOccurrence']", ExcelUtil.GetCellData(row, 3, sheetname), Property_type.XPath);

            if (ExcelUtil.GetCellData(row, 4, sheetname) == "NO")
            {
                Perform.Click(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_PolicyLevelCoverages_ctl_BOP_GeneralInformation_chkBusinessMasterEnhancedEndorsement']", Property_type.XPath);
            }
            Perform.SelectDropDown(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_PolicyLevelCoverages_ctl_BOP_GeneralInformation_ddlBlanketRating']", ExcelUtil.GetCellData(row, 5, sheetname), Property_type.XPath);
            //Policy Level Coverages
            //Check if Addtional Insured data is given
            if (ExcelUtil.GetCellData(row, 6, sheetname) == "YES")
            {
                Perform.Click(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_PolicyLevelCoverages_ctl_BOP_Coverages_chkAdditionalInsured']", Property_type.XPath);
                Perform.SelectDropDown(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_PolicyLevelCoverages_ctl_BOP_Coverages_ddlNumberOfAddlInsureds']", ExcelUtil.GetCellData(row, 7, sheetname), Property_type.XPath);

                if (ExcelUtil.GetCellData(row, 8, sheetname) == "YES")
                    Perform.Click(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_PolicyLevelCoverages_ctl_BOP_Coverages_chkAI_TownhouseAssociates']", Property_type.XPath);
                if (ExcelUtil.GetCellData(row, 9, sheetname) == "YES")
                    Perform.Click(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_PolicyLevelCoverages_ctl_BOP_Coverages_chkAI_EngineersArchitectsSurveyors']", Property_type.XPath);
                if (ExcelUtil.GetCellData(row, 10, sheetname) == "YES")
                    Perform.Click(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_PolicyLevelCoverages_ctl_BOP_Coverages_chkOwnersLesseesContractorsAutomatic']", Property_type.XPath);
                if (ExcelUtil.GetCellData(row, 11, sheetname) == "YES")
                    Perform.Click(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_PolicyLevelCoverages_ctl_BOP_Coverages_chkWaiverOfSubrogation']", Property_type.XPath);
                if (ExcelUtil.GetCellData(row, 12, sheetname) == "YES")
                    Perform.Click(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_PolicyLevelCoverages_ctl_BOP_Coverages_chkOwnersLesseesContractorsWithAddlInsuredReq']", Property_type.XPath);
                if (ExcelUtil.GetCellData(row, 13, sheetname) == "YES")
                    Perform.Click(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_PolicyLevelCoverages_ctl_BOP_Coverages_chkOwnersLesseesContractorsCompletedOps']", Property_type.XPath);

            }
            Console.WriteLine("Additonal Info given");
            
            //Check if Employee Benefits Liability Data is given
            if (ExcelUtil.GetCellData(row, 14, sheetname) == "YES")
            {
                Perform.Click(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_PolicyLevelCoverages_ctl_BOP_Coverages_chkEmployeeBenefitsLiability']", Property_type.XPath);
                Perform.EnterText(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_PolicyLevelCoverages_ctl_BOP_Coverages_txtEBLNumberOfEmployees']", ExcelUtil.GetCellData(row, 15, sheetname), Property_type.XPath);
            }
            Console.WriteLine("Employee benefits Info given");
            //Check for Employment Practices Liability - Claims-Made Basis details
            if (ExcelUtil.GetCellData(row, 16, sheetname) == "NO")
            {
                Perform.Click(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_PolicyLevelCoverages_ctl_BOP_Coverages_chkEPLI']", Property_type.XPath);
                Perform.driver.SwitchTo().Alert().Accept();
            }
            Console.WriteLine("Employee Practices Info given");
            //Check if Contractors Equipment/Installation details are given
            if (ExcelUtil.GetCellData(row, 17, sheetname) == "YES")
            {
                Perform.Click(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_PolicyLevelCoverages_ctl_BOP_Coverages_chkContractorsEquipmentInstallation']", Property_type.XPath);
                if (ExcelUtil.GetCellData(row, 18, sheetname) != "")
                {
                    Perform.SelectDropDown(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_PolicyLevelCoverages_ctl_BOP_Coverages_ddlContractorsPropertyLimitAtEachCoveredJobsite']", ExcelUtil.GetCellData(row, 18, sheetname), Property_type.XPath);
                }
                if (ExcelUtil.GetCellData(row, 19, sheetname) != "")
                {
                    Perform.EnterText(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_PolicyLevelCoverages_ctl_BOP_Coverages_txtContractorsToolsAndEquipmentBlanketLimit']", ExcelUtil.GetCellData(row, 19, sheetname), Property_type.XPath);
                }
                if (ExcelUtil.GetCellData(row, 20, sheetname) != "")
                {
                    Perform.SelectDropDown(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_PolicyLevelCoverages_ctl_BOP_Coverages_ddlContractorsToolsAndEquipmentBlanketSubLimit']", ExcelUtil.GetCellData(row, 20, sheetname), Property_type.XPath);
                }
                if (ExcelUtil.GetCellData(row, 21, sheetname) != "")
                {
                    Perform.EnterText(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_PolicyLevelCoverages_ctl_BOP_Coverages_txtContractorsToolsAndEquipmentScheduledLimit']", ExcelUtil.GetCellData(row, 21, sheetname), Property_type.XPath);
                }
                if (ExcelUtil.GetCellData(row, 22, sheetname) != "")
                {
                    Perform.Click(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_PolicyLevelCoverages_ctl_BOP_Coverages_chkEarthquake']", Property_type.XPath);
                }
                if (ExcelUtil.GetCellData(row, 23, sheetname) != "")
                {
                    Perform.EnterText(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_PolicyLevelCoverages_ctl_BOP_Coverages_txtContractorsRentedLeasedToolsAndEquipmentLimit']", ExcelUtil.GetCellData(row, 23, sheetname), Property_type.XPath);
                }
                if (ExcelUtil.GetCellData(row, 24, sheetname) != "")
                {
                    Perform.EnterText(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_PolicyLevelCoverages_ctl_BOP_Coverages_txtContractorsEmployeesToolsLimit']", ExcelUtil.GetCellData(row, 24, sheetname), Property_type.XPath);
                }
            }
            //check if crime details are given
            if (ExcelUtil.GetCellData(row, 25, sheetname) == "YES")
            {
                Perform.Click(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_PolicyLevelCoverages_ctl_BOP_Coverages_chkCrime']", Property_type.XPath);
                Perform.EnterText(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_PolicyLevelCoverages_ctl_BOP_Coverages_txtCrimeNumberOfEmployees']", ExcelUtil.GetCellData(row, 26, sheetname), Property_type.XPath);
                Perform.EnterText(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_PolicyLevelCoverages_ctl_BOP_Coverages_txtCrimeNumberOfLocations']", ExcelUtil.GetCellData(row, 27, sheetname), Property_type.XPath);
                Perform.SelectDropDown(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_PolicyLevelCoverages_ctl_BOP_Coverages_ddlCrimeTotalLimit']", ExcelUtil.GetCellData(row, 28, sheetname), Property_type.XPath);
            }
            //check if earthquake is selected
            if (ExcelUtil.GetCellData(row, 29, sheetname) == "YES")
            {
                Perform.Click(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_PolicyLevelCoverages_ctl_BOP_Coverages_chkEarthquake']", Property_type.XPath);
            }
            //check if Hired-Auto is selected
            if (ExcelUtil.GetCellData(row, 30, sheetname) == "YES")
            {
                Perform.Click(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_PolicyLevelCoverages_ctl_BOP_Coverages_chkHiredAuto']", Property_type.XPath);
            }
            //check if Non-owned is selected
            if (ExcelUtil.GetCellData(row, 31, sheetname) == "YES")
            {
                Perform.Click(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_PolicyLevelCoverages_ctl_BOP_Coverages_chkNonOwned']", Property_type.XPath);
            }
            //check if electronic data is selected
            if (ExcelUtil.GetCellData(row, 32, sheetname) == "YES")
            {
                Perform.Click(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_PolicyLevelCoverages_ctl_BOP_Coverages_chkElectronicData']", Property_type.XPath);
                Perform.EnterText(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_PolicyLevelCoverages_ctl_BOP_Coverages_txtElectronicDataLimit']", ExcelUtil.GetCellData(row, 32, sheetname), Property_type.XPath);
            }
            Perform.ScreenShot(savescreenshot + "PolicyCoverage.png");
            Perform.test.Log(LogStatus.Info, " Policy coverage Information Entered");
            Perform.driver.FindElement(By.XPath(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_PolicyLevelCoverages_ctl_BOP_Coverages_btnLocations']")).SendKeys(Keys.PageDown);
            Perform.Click(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_PolicyLevelCoverages_ctl_BOP_Coverages_btnLocations']", Property_type.XPath);
            Perform.test.Log(LogStatus.Info, "Proceed to Location");
            Perform.IsElementPresent(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_LocationList_btnSaveAndRate']");
        }
        [Test]
        public void Location(int row, string sheetname, string savescreenshot)
        {
            
            PolicyCoverage(row, "Policy Level Coverages", savescreenshot);
            Perform.driver.Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(10);
            Perform.test = Perform.report.StartTest("Location Page");
            Perform.test.Log(LogStatus.Info, " Enter Location Information");
            String noflocation = ExcelUtil.GetCellData(row, 3, "TestCase");
            //String nofbuilding = ExcelUtil.GetCellData(row, 3, "TestCase");
            //Add Location Information
            PageUtility.LocationCoverage(row, noflocation, sheetname);
            Perform.ScreenShot(savescreenshot + "Location.png");
            Perform.test.Log(LogStatus.Info, "Location Details Entered");
            //Click Rate this Quote button
            Perform.test.Log(LogStatus.Info, "Proceed to Location");
            Perform.driver.FindElement(By.XPath(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_LocationList_btnSaveAndRate']")).SendKeys(Keys.PageDown);
            Perform.Click(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_LocationList_btnSaveAndRate']", Property_type.XPath);
            Perform.IsElementPresent(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_Quote_Summary_ctlQuoteSummaryActions_btnCommContinueToApplication']");

        }
        [Test]
        public void QuoteSummary(int row, string sheetname, string savescreenshot)
        {
            Perform.driver.Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(30);
            Location(row, "Location_0", savescreenshot);
           Perform.test = Perform.report.StartTest("Quote Summary Page");
            Perform.driver.Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(20);
            Perform.waitTillElementToAppear(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_Quote_Summary_ctlQuoteSummaryActions_btnCommIRPM']");
            //Console.WriteLine(ExcelUtil.GetCellData(row, 0, "IRPM"));
            if (ExcelUtil.GetCellData(row, 0, "IRPM") == "IRPM")
            {
                Perform.test.Log(LogStatus.Info, "IRPM is calculated");
                Perform.Click(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_Quote_Summary_ctlQuoteSummaryActions_btnCommIRPM']", Property_type.XPath);
                PageUtility.IRPMValue(row, "IRPM");
                Perform.ScreenShot(savescreenshot + "IRPM.png");
                Perform.driver.FindElement(By.XPath(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_IRPM_btnSubmitRate']")).SendKeys(Keys.PageDown);
                Perform.Click(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_IRPM_btnSubmitRate']", Property_type.XPath);
            }

            Perform.ScreenShot(savescreenshot + "QuoteSummary.png");
            Perform.test.Log(LogStatus.Info, "Validate Quote Summary and proceed to Underwriting");
            Perform.driver.FindElement(By.XPath(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_Quote_Summary_ctlQuoteSummaryActions_btnCommContinueToApplication']")).SendKeys(Keys.PageDown);
            Perform.Click(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_Quote_Summary_ctlQuoteSummaryActions_btnCommContinueToApplication']", Property_type.XPath);
            Perform.IsElementPresent(".//*[@id='cphMain_ctl_WorkflowManager_App_BOP_ctlCommercialUWQuestionList_btnGoToApp']");

        }
        [Test]
        public void Underwriting(int row, string sheetname, string savescreenshot)
        {
            QuoteSummary(row, "", savescreenshot);
            Perform.driver.Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(10);
            Perform.test = Perform.report.StartTest("Policy Underwriting Page");
            //Applicant Information
            Perform.test.Log(LogStatus.Info, "Answer Underwriting");
            int col = 0;
            for (int i = 0; i < 14; i++)
            {
                 if (i == 12)
                 {
                     Perform.driver.FindElement(By.XPath(".//*[@id='cphMain_ctl_WorkflowManager_App_BOP_ctlCommercialUWQuestionList_rptUWQ_ctlCommercialUWQuestionItem_0_rptUWQ_0_rbNo_11']")).SendKeys(Keys.PageDown);
                    System.Threading.Thread.Sleep(500);
                    // Perform.Wait();
                   // Perform.waitTillElementToAppear(".//*[@id='cphMain_ctl_WorkflowManager_App_BOP_ctlCommercialUWQuestionList_rptUWQ_ctlCommercialUWQuestionItem_0_rptUWQ_0_rbNo_" + i + "']");
                }
              
                if (ExcelUtil.GetCellData(row, col, sheetname) == "NO")
                {
                    Perform.Click(".//*[@id='cphMain_ctl_WorkflowManager_App_BOP_ctlCommercialUWQuestionList_rptUWQ_ctlCommercialUWQuestionItem_0_rptUWQ_0_rbNo_" + i + "']", Property_type.XPath);
                    //Console.WriteLine(ExcelUtil.GetCellData(row, col, sheetname));
                }
                if (ExcelUtil.GetCellData(row, col, sheetname) == "YES")
                {
                    Perform.Click(".//*[@id='cphMain_ctl_WorkflowManager_App_BOP_ctlCommercialUWQuestionList_rptUWQ_ctlCommercialUWQuestionItem_0_rptUWQ_0_rbYes_" + i + "']", Property_type.XPath);

                    Perform.waitTillElementToAppear(".//*[@id='cphMain_ctl_WorkflowManager_App_BOP_ctlCommercialUWQuestionList_rptUWQ_ctlCommercialUWQuestionItem_0_rptUWQ_0_txtUWQDescription_" + i + "']");
                    Perform.EnterText(".//*[@id='cphMain_ctl_WorkflowManager_App_BOP_ctlCommercialUWQuestionList_rptUWQ_ctlCommercialUWQuestionItem_0_rptUWQ_0_txtUWQDescription_" + i + "']", ExcelUtil.GetCellData(row, col + 1, sheetname), Property_type.XPath);
                    
                    //Console.WriteLine(ExcelUtil.GetCellData(row, col, sheetname));
                }
                col = col + 2;
            }
            //Business Owner General Info
           // Perform.driver.FindElement(By.XPath(".//*[@id='cphMain_ctl_WorkflowManager_App_BOP_ctlCommercialUWQuestionList_rptUWQ_ctlCommercialUWQuestionItem_1_rptUWQ_1_rbNo_0")).SendKeys(Keys.PageDown);
            int bcol = 28;
            for (int j = 0; j < 9; j++)
            {
              /*  if (j == 2)
                {
                    Perform.driver.FindElement(By.XPath(".//*[@id='cphMain_ctl_WorkflowManager_App_BOP_ctlCommercialUWQuestionList_rptUWQ_ctlCommercialUWQuestionItem_1_rptUWQ_1_rbNo_+j+']")).SendKeys(Keys.PageDown);
                    System.Threading.Thread.Sleep(500);
                }*/
                if (ExcelUtil.GetCellData(row, bcol, sheetname) == "NO")
                {
                    Perform.Click(".//*[@id='cphMain_ctl_WorkflowManager_App_BOP_ctlCommercialUWQuestionList_rptUWQ_ctlCommercialUWQuestionItem_1_rptUWQ_1_rbNo_" + j + "']", Property_type.XPath);
                }
                if (ExcelUtil.GetCellData(row, bcol, sheetname) == "YES")
                {
                    Perform.Click(".//*[@id='cphMain_ctl_WorkflowManager_App_BOP_ctlCommercialUWQuestionList_rptUWQ_ctlCommercialUWQuestionItem_1_rptUWQ_1_rbYes_" + j + "']", Property_type.XPath);
                    // Console.WriteLine("Clicked" + bcol);
                    Perform.waitTillElementToAppear(".//*[@id='cphMain_ctl_WorkflowManager_App_BOP_ctlCommercialUWQuestionList_rptUWQ_ctlCommercialUWQuestionItem_1_rptUWQ_1_txtUWQDescription_" + j + "']");
                    Perform.EnterText(".//*[@id='cphMain_ctl_WorkflowManager_App_BOP_ctlCommercialUWQuestionList_rptUWQ_ctlCommercialUWQuestionItem_1_rptUWQ_1_txtUWQDescription_" + j + "']", ExcelUtil.GetCellData(row, bcol + 1, sheetname), Property_type.XPath);
                }
                bcol = bcol + 2;
            }
            int premcol = 46;
            Perform.driver.FindElement(By.XPath(".//*[@id='cphMain_ctl_WorkflowManager_App_BOP_ctlCommercialUWQuestionList_rptUWQ_ctlCommercialUWQuestionItem_2_rptUWQ_2_rbNo_0']")).SendKeys(Keys.PageDown);
            System.Threading.Thread.Sleep(500);
            for (int z = 0; z < 5; z++)
            {
                if (ExcelUtil.GetCellData(row, premcol, sheetname) == "NO")
                {
                    Perform.Click(".//*[@id='cphMain_ctl_WorkflowManager_App_BOP_ctlCommercialUWQuestionList_rptUWQ_ctlCommercialUWQuestionItem_2_rptUWQ_2_rbNo_" + z + "']", Property_type.XPath);
                    //Console.WriteLine("Clicked" + premcol);
                }
                if (ExcelUtil.GetCellData(row, premcol, sheetname) == "YES")
                {
                    if (z == 2)
                    {
                        Perform.Click(".//*[@id='cphMain_ctl_WorkflowManager_App_BOP_ctlCommercialUWQuestionList_rptUWQ_ctlCommercialUWQuestionItem_2_rptUWQ_2_rbYes_" + z + "']", Property_type.XPath);
                        //Console.WriteLine("Clicked" + premcol);
                    }
                    else
                    {
                        Perform.Click(".//*[@id='cphMain_ctl_WorkflowManager_App_BOP_ctlCommercialUWQuestionList_rptUWQ_ctlCommercialUWQuestionItem_2_rptUWQ_2_rbYes_" + z + "']", Property_type.XPath);
                        // Console.WriteLine("Clicked" + premcol);
                        Perform.EnterText(".//*[@id='cphMain_ctl_WorkflowManager_App_BOP_ctlCommercialUWQuestionList_rptUWQ_ctlCommercialUWQuestionItem_2_rptUWQ_2_txtUWQDescription_" + z + "']", ExcelUtil.GetCellData(row, bcol + 1, sheetname), Property_type.XPath);
                    }
                }
                if (z == 2)
                {
                    premcol = premcol + 1;
                }

                else
                {
                    premcol = premcol + 2;
                }
            }
            if (ExcelUtil.GetCellData(row, 55, sheetname) != "")
            {
                int apt = 55;
                for (int a = 0; a < 4; a++)
                {
                    if (ExcelUtil.GetCellData(row, apt, sheetname) == "NO")
                    {
                        Perform.Click(".//*[@id='cphMain_ctl_WorkflowManager_App_BOP_ctlCommercialUWQuestionList_rptUWQ_ctlCommercialUWQuestionItem_3_rptUWQ_3_rbNo_" + a + "']", Property_type.XPath);
                        //Console.WriteLine("Clicked" + apt);
                    }
                    if (ExcelUtil.GetCellData(row, apt, sheetname) == "YES")
                    {
                        if (a == 2)
                        {
                            Perform.Click(".//*[@id='cphMain_ctl_WorkflowManager_App_BOP_ctlCommercialUWQuestionList_rptUWQ_ctlCommercialUWQuestionItem_3_rptUWQ_3_rbYes_" + a + "']", Property_type.XPath);
                            //Console.WriteLine("Clicked" + apt);
                        }
                        else
                        {
                            Perform.Click(".//*[@id='cphMain_ctl_WorkflowManager_App_BOP_ctlCommercialUWQuestionList_rptUWQ_ctlCommercialUWQuestionItem_3_rptUWQ_3_rbYes_" + a + "']", Property_type.XPath);
                            Perform.EnterText(".//*[@id='cphMain_ctl_WorkflowManager_App_BOP_ctlCommercialUWQuestionList_rptUWQ_ctlCommercialUWQuestionItem_3_rptUWQ_3_txtUWQDescription_" + a + "']", ExcelUtil.GetCellData(row, apt + 1, sheetname), Property_type.XPath);
                            //Console.WriteLine("Clicked" + apt);

                        }
                    }

                    if (a == 2)
                    {
                        apt = apt + 1;
                    }
                    else
                    {
                        apt = apt + 2;
                    }
                }
            }
            Perform.test.Log(LogStatus.Info, "Underwriting Questions Answered");
            Perform.ScreenShot(savescreenshot + "Underwriting.png");
            Perform.test.Log(LogStatus.Info, "Proceed to Application");
            //Perform.driver.FindElement(By.XPath(".//*[@id='cphMain_ctl_WorkflowManager_App_BOP_ctlCommercialUWQuestionList_btnGoToApp']")).SendKeys(Keys.PageDown);
            Perform.Click(".//*[@id='cphMain_ctl_WorkflowManager_App_BOP_ctlCommercialUWQuestionList_btnGoToApp']", Property_type.XPath);
            Perform.IsElementPresent(".//*[@id='btnShowEffectiveDate']");
           
        }
        [Test]
        public void Application(int row, string sheetname, string savescreenshot)
        {
            Underwriting(row, "Underwriting Questions", savescreenshot);
           Perform.test = Perform.report.StartTest("Application Page");
            string addapp = ExcelUtil.GetCellData(row, 13, "TestCase");
            Perform.test.Log(LogStatus.Info, "Enter Application Information");
            PageUtility.Additonal_App(row, addapp, sheetname);


            Perform.ScreenShot(savescreenshot + "Application.png");
            Perform.IsElementPresent(".//*[@id='cphMain_ctl_WorkflowManager_App_BOP_ctl_AppSection_BOP_ctl_Billing_Info_PPA_ddMethod']");


        }



        [Test]
        public void BillingInfo(int row, string sheetname, string savescreenshot)
        {
            Application(row, "Application_Location_0", savescreenshot);
          Perform.test = Perform.report.StartTest("Billing Info Page");
            Perform.test.Log(LogStatus.Info, "Enter Billing Info");
            string nofitems = ExcelUtil.GetCellData(row, 0, "Contractor");
            if (ExcelUtil.GetCellData(row, 0, "Contractor") != "")
            {

                PageUtility.ContractorInfo(row, nofitems, sheetname);
            }
            string nofinsured = ExcelUtil.GetCellData(row, 0, "Additional Insured");
            if (ExcelUtil.GetCellData(row, 0, "Additional Insured") != "")
            {

                PageUtility.Additional_Insured(row, nofinsured, sheetname);
            }
            Perform.SelectDropDown(".//*[@id='cphMain_ctl_WorkflowManager_App_BOP_ctl_AppSection_BOP_ctl_Billing_Info_PPA_ddMethod']", ExcelUtil.GetCellData(row, 0, sheetname), Property_type.XPath);
            Perform.SelectDropDown(".//*[@id='cphMain_ctl_WorkflowManager_App_BOP_ctl_AppSection_BOP_ctl_Billing_Info_PPA_ddPayPlan']", ExcelUtil.GetCellData(row, 1, sheetname), Property_type.XPath);
            Perform.SelectDropDown(".//*[@id='cphMain_ctl_WorkflowManager_App_BOP_ctl_AppSection_BOP_ctl_Billing_Info_PPA_ddBillTo']", ExcelUtil.GetCellData(row, 2, sheetname), Property_type.XPath);
            Perform.Click(".//*[@id='btnShowEffectiveDate']", Property_type.XPath);
            Perform.Click(".//*[@id='txtEffectiveDate']", Property_type.XPath);
            //Select Effective Date
            /*  Perform.waitTillElementToAppear(".//*[@id='ui-datepicker-div']/div[1]/div/select[1]");
              Perform.SelectDropDown(".//*[@id='ui-datepicker-div']/div[1]/div/select[1]", ExcelUtil.GetCellData(row, 3, sheetname), Property_type.XPath);
              Perform.waitTillElementToAppear(".//*[@id='ui-datepicker-div']/div[1]/div/select[2]");
              Perform.SelectDropDown(".//*[@id='ui-datepicker-div']/div[1]/div/select[2]", ExcelUtil.GetCellData(row, 5, sheetname), Property_type.XPath);
              Perform.waitTillElementToAppear(ExcelUtil.GetCellData(row, 4, sheetname));
              Perform.Click(ExcelUtil.GetCellData(row, 4, sheetname), Property_type.LinkText);*/
            Perform.EnterText(".//*[@id='txtEffectiveDate']", ExcelUtil.GetCellData(row, 6, sheetname),Property_type.XPath);
            Perform.driver.FindElement(By.XPath(".//*[@id='txtEffectiveDate']")).SendKeys(Keys.Enter);
            System.Threading.Thread.Sleep(500);
            Perform.driver.FindElement(By.Id("btnEffectiveDateDone")).SendKeys(Keys.Enter);
            Perform.test.Log(LogStatus.Info, "The Quote is rated");
            System.Threading.Thread.Sleep(500);
            Perform.IsElementPresent(".//*[@id='cphMain_ctl_WorkflowManager_App_BOP_ctl_BOP_QuoteSummary_ctlQuoteSummaryActions_btnContinueToApp']");
        }
        [Test]
        public void FinalizePage()
        {
            int sheetrownum = ExcelUtil.getRowCount("TestCase");
            try
            {

                for (int i = 5; i < sheetrownum; i++)
                {


                    Console.WriteLine("Processing Row " + i);
                    System.IO.Directory.CreateDirectory(path + ExcelUtil.GetCellData(i, 0, "TestCase"));
                    string savescreenshot = path + ExcelUtil.GetCellData(i, 0, "TestCase") + "\\";
                    Console.WriteLine(savescreenshot);
                    // PolicyCoverage(i, "PolicyCoverage");
                    BillingInfo(i, "Billing Info", savescreenshot);
                   
                    Perform.test = Perform.report.StartTest("Velocirater Commercial BOP");
                    Perform.Click(".//*[@id='cphMain_ctl_WorkflowManager_App_BOP_ctl_BOP_QuoteSummary_ctlQuoteSummaryActions_btnContinueToApp']", Property_type.XPath);
                    Console.WriteLine("Application Finalized");
                    Perform.test.Log(LogStatus.Info, "Application Finalized");
                    Perform.driver.Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(30);
                    Perform.waitTillElementToAppear(".//*[@id='CrumbsLogoutLink']");
                    Perform.CheckTitle("Make a Payment");
                    Perform.ScreenShot(savescreenshot + "Policy" + (i - 4) + "_" + DateTime.Now.ToString("yyyyMMddhhmmss") + ".png");
                  
                    Perform.Click(".//*[@id='CrumbsLogoutLink']", Property_type.XPath);
                    Perform.test.Log(LogStatus.Info, "Logged out");
                    Console.WriteLine("Policy" + (i - 4) + " issued");
                }

            }
            catch (Exception e)
            {

                Console.WriteLine("Error:" + e);
                Assert.True(false);

            }
                
            }


        


        [TearDown]
        public void Cleanup()
        {

            if (TestContext.CurrentContext.Result.Outcome == ResultState.Failure)

                Perform.ScreenShot(path + "Failure.png");
            Perform.driver.Close();
            Perform.report.EndTest(Perform.test);
            Perform.report.Flush();
        }
    }
}
