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

namespace VR_Personal_Auto
{
    public class VR_NB_PPA
    {
        String url;
        String path = @"C:\Users\imnay\Documents\Visual Studio 2015\Projects\VR_Personal_Auto\VR_Personal_Auto\Output\";
        String ReportPath = @"C:\Users\imnay\Documents\Visual Studio 2015\Projects\VR_Personal_Auto\VR_Personal_Auto\Report\";


        [SetUp]
        public void Initialize()

        {
            Perform.Browser("chrome");
           url = "http://www.ifmig.net/NewPublicSite/NewPublicHome.aspx";
            ExcelUtil.setExcelFile(@"C:\Users\imnay\Documents\Visual Studio 2015\Projects\VR_Personal_Auto\VR_Personal_Auto\Excel\Personal_Auto.xlsx");
            Perform.report= new ExtentReports(ReportPath + "Report.html", CultureInfo.GetCultureInfo("es-ES"), true, DisplayOrder.NewestFirst);
            Perform.report.LoadConfig(ReportPath + "extent-config.xml");
        }
        [Test]
        public void LoginPage()
        {
            Perform.test = Perform.report.StartTest("Login Page");
            //Login Page
            Property_Collection.driver.Navigate().GoToUrl(url);
            Console.WriteLine("Browser Opened");
            Perform.test.Log(LogStatus.Info, "Browser Opened");
            //Click on Agents Only Link
            Perform.Click(".//*[@id='Footer_AgentsOnlyLink']", Property_type.XPath);

            //Page 1
            //Enter Username
            Perform.EnterText("Application_txtUsername", "DonBrewtonTest", Property_type.Id);
            //Enter Password
            Perform.EnterText("Application_txtPassword", "DonBrewtonTest1", Property_type.Id);
            Perform.test.Log(LogStatus.Info, "User Credentials entered");
            Perform.Click("Application_btnLogin", Property_type.Id);
            Perform.test.Log(LogStatus.Info, "Login button Clicked");
            Perform.CheckTitle("Welcome Agents!");
            Console.WriteLine("User Details Entered");
          

        }
        [Test]
        public void Menu()
        {
            LoginPage();
            Perform.test = Perform.report.StartTest("Menu Page");
            //Page 2
            Perform.waitTillElementToAppear(".//*[@id='ulaitem0_0']");
            Perform.Click(".//*[@id='ulaitem0_0']", Property_type.XPath);
            System.Threading.Thread.Sleep(2000);
            Perform.waitTillElementToAppear("//*[@id='ulaitem0_0_1']");
            Perform.Click("//*[@id='ulaitem0_0_1']", Property_type.XPath);
            Console.WriteLine("VelociRater is Clicked");
            Perform.test.Log(LogStatus.Info, "Velocirater Clicked");
            Perform.CheckTitle("VelociRater");

            //Page 3
            Perform.Click(".//*[@id='main']/table/tbody/tr/td[2]/div[1]/div[2]/div/input", Property_type.XPath);
            Perform.test.Log(LogStatus.Info, "New Personal Auto is Clicked");
            Console.WriteLine(Perform.GetText(".//*[@id='main']/table/tbody/tr/td[2]/div[1]/div[2]/div/input", Property_type.XPath) + " is selected");
            Perform.PageContains("Underwriting Questions");
        }
        [Test]
        public void PopUpPage(int row, String Sheetname, string savescreenshot)
        {
            Menu();
            Perform.test = Perform.report.StartTest("General Underwriting Page");
            //UnderWriting Questions
            Perform.click_on_webElements("//input[contains(@id,'radNo_')]", Property_Collection.driver);
            Perform.Click("radMultiPolicyNo", Property_type.Id);
            Console.WriteLine("Underwriting questions answered");
            Perform.test.Log(LogStatus.Info, "Underwriting questions answered as NO");
            Perform.Click("cphMain_ctlUWQuestionsPopup_btnSave", Property_type.Id);
            Perform.PageContains("*First Name");

        }




        [Test]
        public void PolicyHolder(int row, String Sheetname, string savescreenshot)
        {


            PopUpPage(row, " ", savescreenshot);
            Perform.test = Perform.report.StartTest("PolicyHolder Page");
            //PolicyHolder Page
            //PolicyHolder1 Details
            Perform.EnterText(".//*[@id='cphMain_ctl_Master_Edit_ctlIsuredList_ctlInsured_txtFirstName']", ExcelUtil.GetCellData(row, 0, Sheetname), Property_type.XPath);
            Perform.EnterText(".//*[@id='cphMain_ctl_Master_Edit_ctlIsuredList_ctlInsured_txtMiddleName']", ExcelUtil.GetCellData(row, 1, Sheetname), Property_type.XPath);
            Perform.EnterText(".//*[@id='cphMain_ctl_Master_Edit_ctlIsuredList_ctlInsured_txtLastName']", ExcelUtil.GetCellData(row, 2, Sheetname), Property_type.XPath);
            Perform.SelectDropDown(".//*[@id='cphMain_ctl_Master_Edit_ctlIsuredList_ctlInsured_ddSuffix']", ExcelUtil.GetCellData(row, 3, Sheetname), Property_type.XPath);
            Perform.SelectDropDown(".//*[@id='cphMain_ctl_Master_Edit_ctlIsuredList_ctlInsured_ddSex']", ExcelUtil.GetCellData(row, 4, Sheetname), Property_type.XPath);
            Perform.EnterText(".//*[@id='cphMain_ctl_Master_Edit_ctlIsuredList_ctlInsured_txtSSN']", ExcelUtil.GetCellData(row, 5, Sheetname), Property_type.XPath);
            Perform.EnterText(".//*[@id='cphMain_ctl_Master_Edit_ctlIsuredList_ctlInsured_txtBirthDate']", ExcelUtil.GetCellData(row, 6, Sheetname), Property_type.XPath);
            Perform.EnterText(".//*[@id='cphMain_ctl_Master_Edit_ctlIsuredList_ctlInsured_txtEmail']", ExcelUtil.GetCellData(row, 7, Sheetname), Property_type.XPath);
            Perform.EnterText(".//*[@id='cphMain_ctl_Master_Edit_ctlIsuredList_ctlInsured_txtPhone']", ExcelUtil.GetCellData(row, 8, Sheetname), Property_type.XPath);
            Perform.EnterText(".//*[@id='cphMain_ctl_Master_Edit_ctlIsuredList_ctlInsured_txtPhoneExt']", ExcelUtil.GetCellData(row, 9, Sheetname), Property_type.XPath);
            Perform.SelectDropDown(".//*[@id='cphMain_ctl_Master_Edit_ctlIsuredList_ctlInsured_ddPhoneType']", ExcelUtil.GetCellData(row, 10, Sheetname), Property_type.XPath);
            Perform.EnterText(".//*[@id='cphMain_ctl_Master_Edit_ctlIsuredList_ctlInsured_txtStreetNum']", ExcelUtil.GetCellData(row, 11, Sheetname), Property_type.XPath);
            Perform.EnterText(".//*[@id='cphMain_ctl_Master_Edit_ctlIsuredList_ctlInsured_txtStreetName']", ExcelUtil.GetCellData(row, 12, Sheetname), Property_type.XPath);
            Perform.EnterText(".//*[@id='cphMain_ctl_Master_Edit_ctlIsuredList_ctlInsured_txtAptNum']", ExcelUtil.GetCellData(row, 13, Sheetname), Property_type.XPath);
            Perform.EnterText(".//*[@id='cphMain_ctl_Master_Edit_ctlIsuredList_ctlInsured_txtPOBox']", ExcelUtil.GetCellData(row, 14, Sheetname), Property_type.XPath);
            Perform.EnterText(".//*[@id='cphMain_ctl_Master_Edit_ctlIsuredList_ctlInsured_txtZipCode']", ExcelUtil.GetCellData(row, 15, Sheetname), Property_type.XPath);
            //Perform.Wait();

            Perform.EnterText(".//*[@id='cphMain_ctl_Master_Edit_ctlIsuredList_ctlInsured_txtCityName']", ExcelUtil.GetCellData(row, 16, Sheetname), Property_type.XPath);
            Perform.SelectDropDown(".//*[@id='cphMain_ctl_Master_Edit_ctlIsuredList_ctlInsured_ddStateAbbrev']", ExcelUtil.GetCellData(row, 17, Sheetname), Property_type.XPath);
            Perform.EnterText(".//*[@id='cphMain_ctl_Master_Edit_ctlIsuredList_ctlInsured_txtGaragedCounty']", ExcelUtil.GetCellData(row, 18, Sheetname), Property_type.XPath);
            Console.WriteLine("Policyholder1 details are entered");

            //PolicyHolder2 Details
            Perform.Click(".//*[@id='cphMain_ctl_Master_Edit_ctlIsuredList_ctlInsured1_lblInsuredTitle']", Property_type.XPath);
            Perform.waitTillElementToAppear(".//*[@id='cphMain_ctl_Master_Edit_ctlIsuredList_ctlInsured1_txtFirstName']");
            Perform.EnterText(".//*[@id='cphMain_ctl_Master_Edit_ctlIsuredList_ctlInsured1_txtFirstName']", ExcelUtil.GetCellData(row, 19, Sheetname), Property_type.XPath);
            Perform.EnterText(".//*[@id='cphMain_ctl_Master_Edit_ctlIsuredList_ctlInsured1_txtMiddleName']", ExcelUtil.GetCellData(row, 20, Sheetname), Property_type.XPath);
            Perform.EnterText(".//*[@id='cphMain_ctl_Master_Edit_ctlIsuredList_ctlInsured1_txtLastName']", ExcelUtil.GetCellData(row, 21, Sheetname), Property_type.XPath);
            Perform.SelectDropDown(".//*[@id='cphMain_ctl_Master_Edit_ctlIsuredList_ctlInsured1_ddSuffix']", ExcelUtil.GetCellData(row, 22, Sheetname), Property_type.XPath);
            Perform.SelectDropDown(".//*[@id='cphMain_ctl_Master_Edit_ctlIsuredList_ctlInsured1_ddSex']", ExcelUtil.GetCellData(row, 23, Sheetname), Property_type.XPath);
            Perform.EnterText(".//*[@id='cphMain_ctl_Master_Edit_ctlIsuredList_ctlInsured1_txtSSN']", ExcelUtil.GetCellData(row, 24, Sheetname), Property_type.XPath);
            Perform.EnterText(".//*[@id='cphMain_ctl_Master_Edit_ctlIsuredList_ctlInsured1_txtBirthDate']", ExcelUtil.GetCellData(row, 25, Sheetname), Property_type.XPath);
            Perform.EnterText(".//*[@id='cphMain_ctl_Master_Edit_ctlIsuredList_ctlInsured1_txtEmail']", ExcelUtil.GetCellData(row, 26, Sheetname), Property_type.XPath);
            Perform.EnterText(".//*[@id='cphMain_ctl_Master_Edit_ctlIsuredList_ctlInsured1_txtPhone']", ExcelUtil.GetCellData(row, 27, Sheetname), Property_type.XPath);
            Perform.EnterText(".//*[@id='cphMain_ctl_Master_Edit_ctlIsuredList_ctlInsured1_txtPhoneExt']", ExcelUtil.GetCellData(row, 28, Sheetname), Property_type.XPath);
            Console.WriteLine("Policy holder 2 details are entered");
            //Take Screenshot
            Perform.ScreenShot(savescreenshot + "PolicyHolderPage.png");
            Perform.test.Log(LogStatus.Info, "Policyholder details are entered");
            Perform.Click(".//*[@id='cphMain_ctl_Master_Edit_ctlIsuredList_btnSaveAndGotoDrivers']", Property_type.XPath);
            Perform.test.Log(LogStatus.Info, "Proceed to Driver Details");
            Perform.IsElementPresent(".//*[@id='cphMain_ctl_Master_Edit_ctlDriverList_btnSaveAndGotoVehicles']");
        }

        [Test]
        public void DriverPage(int row, String Sheetname, string savescreenshot)
        {
            PolicyHolder(row, "Policyholder", savescreenshot);
            Perform.test = Perform.report.StartTest("Driver Details Page");
            Perform.test.Log(LogStatus.Info, "Enter Driver Details");
            String nofdriver = ExcelUtil.GetCellData(row, 2, "TestCase");

           PageUtility.AddDriverDetails(row, nofdriver, Sheetname);
            Perform.ScreenShot(savescreenshot + "DriverDetails.png");
            Perform.test.Log(LogStatus.Info, "Driver Details Entered");
            Perform.Click(".//*[@id='cphMain_ctl_Master_Edit_ctlDriverList_btnSaveAndGotoVehicles']", Property_type.XPath);
            Perform.test.Log(LogStatus.Info, "Proceed to Vehicles Page");
            Perform.IsElementPresent(".//*[@id='cphMain_ctl_Master_Edit_ctlVehicleList_btnSaveandGotoCoverages']");
            /* Perform.Click("//*[@id='cphMain_ctl_Master_Edit_ctlDriverList_bnAddDriver']", Property_type.XPath);
             //Driver1 Details
             string result = ExcelUtil.GetCellData(row, 0, Sheetname);

             if (result != "")
             {
                 Perform.Click(".//*[@id='cphMain_ctl_Master_Edit_ctlDriverList_Repeater1_ctlDriver_PPAControl_0_btnCopyFromPh1_0']", Property_type.XPath);
                 Perform.EnterText(".//*[@id='cphMain_ctl_Master_Edit_ctlDriverList_Repeater1_ctlDriver_PPAControl_0_txtFirstName_0']", ExcelUtil.GetCellData(row, 0, Sheetname), Property_type.XPath);
                 Perform.EnterText(".//*[@id='cphMain_ctl_Master_Edit_ctlDriverList_Repeater1_ctlDriver_PPAControl_0_txtMiddleName_0']", ExcelUtil.GetCellData(row, 1, Sheetname), Property_type.XPath);
                 Perform.EnterText(".//*[@id='cphMain_ctl_Master_Edit_ctlDriverList_Repeater1_ctlDriver_PPAControl_0_txtLastname_0']", ExcelUtil.GetCellData(row, 2, Sheetname), Property_type.XPath);
                 Perform.SelectDropDown(".//*[@id='cphMain_ctl_Master_Edit_ctlDriverList_Repeater1_ctlDriver_PPAControl_0_ddSuffix_0']", ExcelUtil.GetCellData(row, 3, Sheetname), Property_type.XPath);
                 Perform.EnterText(".//*[@id='cphMain_ctl_Master_Edit_ctlDriverList_Repeater1_ctlDriver_PPAControl_0_txtBirthDate_0']", ExcelUtil.GetCellData(row, 4, Sheetname), Property_type.XPath);
                 Perform.SelectDropDown(".//*[@id='cphMain_ctl_Master_Edit_ctlDriverList_Repeater1_ctlDriver_PPAControl_0_ddSex_0']", ExcelUtil.GetCellData(row, 5, Sheetname), Property_type.XPath);
                 Perform.SelectDropDown(".//*[@id='cphMain_ctl_Master_Edit_ctlDriverList_Repeater1_ctlDriver_PPAControl_0_ddMaritialStatus_0']", ExcelUtil.GetCellData(row, 6, Sheetname), Property_type.XPath);
                 Perform.EnterText(".//*[@id='cphMain_ctl_Master_Edit_ctlDriverList_Repeater1_ctlDriver_PPAControl_0_txtDLNumber_0']", ExcelUtil.GetCellData(row, 7, Sheetname), Property_type.XPath);
                 Perform.SelectDropDown(".//*[@id='cphMain_ctl_Master_Edit_ctlDriverList_Repeater1_ctlDriver_PPAControl_0_ddDLState_0']", ExcelUtil.GetCellData(row, 8, Sheetname), Property_type.XPath);
                 Perform.SelectDropDown(".//*[@id='cphMain_ctl_Master_Edit_ctlDriverList_Repeater1_ctlDriver_PPAControl_0_ddRelationToPolicyHolder_0']", ExcelUtil.GetCellData(row, 9, Sheetname), Property_type.XPath);
                 Perform.SelectDropDown(".//*[@id='cphMain_ctl_Master_Edit_ctlDriverList_Repeater1_ctlDriver_PPAControl_0_ddRatedOrExcludedDriver_0']", ExcelUtil.GetCellData(row, 10, Sheetname), Property_type.XPath);
                 Console.WriteLine("Driver1 details Entered");
                 //Check if 2nd driver is present
                 string result1 = ExcelUtil.GetCellData(row, 24, Sheetname);
                 if (result1 == "")
                 {
                     Perform.ScreenShot(@"C:\Users\imnay\Documents\Visual Studio 2015\Projects\ClassLibrary1\ClassLibrary1\Output\DriverDetails.png");
                     Perform.Click(".//*[@id='cphMain_ctl_Master_Edit_ctlDriverList_btnSaveAndGotoVehicles']", Property_type.XPath);
                 }
                 else
                 {
                     Perform.Click(".//*[@id='cphMain_ctl_Master_Edit_ctlDriverList_bnAddDriver']", Property_type.XPath);

                     Perform.EnterText(".//*[@id='cphMain_ctl_Master_Edit_ctlDriverList_Repeater1_ctlDriver_PPAControl_1_txtFirstName_1']", ExcelUtil.GetCellData(row, 24, Sheetname), Property_type.XPath);

                     Perform.EnterText(".//*[@id='cphMain_ctl_Master_Edit_ctlDriverList_Repeater1_ctlDriver_PPAControl_1_txtMiddleName_1']", ExcelUtil.GetCellData(row, 25, Sheetname), Property_type.XPath);
                     Perform.EnterText(".//*[@id='cphMain_ctl_Master_Edit_ctlDriverList_Repeater1_ctlDriver_PPAControl_1_txtLastname_1']", ExcelUtil.GetCellData(row, 26, Sheetname), Property_type.XPath);
                     Perform.SelectDropDown(".//*[@id='cphMain_ctl_Master_Edit_ctlDriverList_Repeater1_ctlDriver_PPAControl_1_ddSuffix_1']", ExcelUtil.GetCellData(row, 27, Sheetname), Property_type.XPath);
                     Perform.EnterText(".//*[@id='cphMain_ctl_Master_Edit_ctlDriverList_Repeater1_ctlDriver_PPAControl_1_txtBirthDate_1']", ExcelUtil.GetCellData(row, 28, Sheetname), Property_type.XPath);

                     Perform.SelectDropDown("//select[contains(@id,'ddSex_1')]", ExcelUtil.GetCellData(row, 29, Sheetname), Property_type.XPath);

                     Perform.SelectDropDown("//select[contains(@id,'ddMaritialStatus_1')]", ExcelUtil.GetCellData(row, 30, Sheetname), Property_type.XPath);
                     Perform.EnterText("//input[contains(@id,'txtDLNumber_1')]", ExcelUtil.GetCellData(row, 31, Sheetname), Property_type.XPath);
                     Perform.SelectDropDown(".//*[@id='cphMain_ctl_Master_Edit_ctlDriverList_Repeater1_ctlDriver_PPAControl_1_ddDLState_1']", ExcelUtil.GetCellData(row, 32, Sheetname), Property_type.XPath);
                     Perform.SelectDropDown(".//*[@id='cphMain_ctl_Master_Edit_ctlDriverList_Repeater1_ctlDriver_PPAControl_1_ddRelationToPolicyHolder_1']", ExcelUtil.GetCellData(row, 33, Sheetname), Property_type.XPath);
                     Perform.SelectDropDown(".//*[@id='cphMain_ctl_Master_Edit_ctlDriverList_Repeater1_ctlDriver_PPAControl_1_ddRatedOrExcludedDriver_1']", ExcelUtil.GetCellData(row, 34, Sheetname), Property_type.XPath);
                     Console.WriteLine("Driver2 details Entered");
                     String result2 = ExcelUtil.GetCellData(row, 48, Sheetname);
                     //Check if 3rd driver is present
                     if (result2 == "")
                     {
                         Perform.ScreenShot(@"C:\Users\imnay\Documents\Visual Studio 2015\Projects\ClassLibrary1\ClassLibrary1\Output\DriverDetails.png");
                         Perform.Click(".//*[@id='cphMain_ctl_Master_Edit_ctlDriverList_btnSaveAndGotoVehicles']", Property_type.XPath);
                     }
                     else
                     {
                         Perform.Click(".//*[@id='cphMain_ctl_Master_Edit_ctlDriverList_bnAddDriver']", Property_type.XPath);

                         Perform.EnterText(".//*[@id='cphMain_ctl_Master_Edit_ctlDriverList_Repeater1_ctlDriver_PPAControl_2_txtFirstName_2']", ExcelUtil.GetCellData(row, 48, Sheetname), Property_type.XPath);

                         Perform.EnterText(".//*[@id='cphMain_ctl_Master_Edit_ctlDriverList_Repeater1_ctlDriver_PPAControl_2_txtMiddleName_2']", ExcelUtil.GetCellData(row, 49, Sheetname), Property_type.XPath);
                         Perform.EnterText(".//*[@id='cphMain_ctl_Master_Edit_ctlDriverList_Repeater1_ctlDriver_PPAControl_2_txtLastname_2']", ExcelUtil.GetCellData(row, 50, Sheetname), Property_type.XPath);
                         Perform.SelectDropDown(".//*[@id='cphMain_ctl_Master_Edit_ctlDriverList_Repeater1_ctlDriver_PPAControl_2_ddSuffix_2']", ExcelUtil.GetCellData(row, 51, Sheetname), Property_type.XPath);
                         Perform.EnterText(".//*[@id='cphMain_ctl_Master_Edit_ctlDriverList_Repeater1_ctlDriver_PPAControl_2_txtBirthDate_2']", ExcelUtil.GetCellData(row, 52, Sheetname), Property_type.XPath);

                         Perform.SelectDropDown("//select[contains(@id,'ddSex_2')]", ExcelUtil.GetCellData(row, 53, Sheetname), Property_type.XPath);

                         Perform.SelectDropDown("//select[contains(@id,'ddMaritialStatus_2')]", ExcelUtil.GetCellData(row, 54, Sheetname), Property_type.XPath);
                         Perform.EnterText("//input[contains(@id,'txtDLNumber_2')]", ExcelUtil.GetCellData(row, 55, Sheetname), Property_type.XPath);
                         Perform.SelectDropDown(".//*[@id='cphMain_ctl_Master_Edit_ctlDriverList_Repeater1_ctlDriver_PPAControl_2_ddDLState_2']", ExcelUtil.GetCellData(row, 56, Sheetname), Property_type.XPath);
                         Perform.SelectDropDown(".//*[@id='cphMain_ctl_Master_Edit_ctlDriverList_Repeater1_ctlDriver_PPAControl_2_ddRelationToPolicyHolder_2']", ExcelUtil.GetCellData(row, 57, Sheetname), Property_type.XPath);
                         Perform.SelectDropDown(".//*[@id='cphMain_ctl_Master_Edit_ctlDriverList_Repeater1_ctlDriver_PPAControl_2_ddRatedOrExcludedDriver_2']", ExcelUtil.GetCellData(row, 58, Sheetname), Property_type.XPath);
                         Console.WriteLine("Driver3 details Entered");
                         String result3 = ExcelUtil.GetCellData(row, 72, Sheetname);
                         if (result3 == "")
                         {
                             Perform.ScreenShot(@"C:\Users\imnay\Documents\Visual Studio 2015\Projects\ClassLibrary1\ClassLibrary1\Output\DriverDetails.png");

                             Perform.Click(".//*[@id='cphMain_ctl_Master_Edit_ctlDriverList_btnSaveAndGotoVehicles']", Property_type.XPath);
                         }
                         else
                         {
                             Perform.Click(".//*[@id='cphMain_ctl_Master_Edit_ctlDriverList_bnAddDriver']", Property_type.XPath);

                             Perform.EnterText(".//*[@id='cphMain_ctl_Master_Edit_ctlDriverList_Repeater1_ctlDriver_PPAControl_3_txtFirstName_3']", ExcelUtil.GetCellData(row, 72, Sheetname), Property_type.XPath);

                             Perform.EnterText(".//*[@id='cphMain_ctl_Master_Edit_ctlDriverList_Repeater1_ctlDriver_PPAControl_3_txtMiddleName_3']", ExcelUtil.GetCellData(row, 73, Sheetname), Property_type.XPath);
                             Perform.EnterText(".//*[@id='cphMain_ctl_Master_Edit_ctlDriverList_Repeater1_ctlDriver_PPAControl_3_txtLastname_3']", ExcelUtil.GetCellData(row, 74, Sheetname), Property_type.XPath);
                             Perform.SelectDropDown(".//*[@id='cphMain_ctl_Master_Edit_ctlDriverList_Repeater1_ctlDriver_PPAControl_3_ddSuffix_3']", ExcelUtil.GetCellData(row, 75, Sheetname), Property_type.XPath);
                             Perform.EnterText(".//*[@id='cphMain_ctl_Master_Edit_ctlDriverList_Repeater1_ctlDriver_PPAControl_3_txtBirthDate_3']", ExcelUtil.GetCellData(row, 76, Sheetname), Property_type.XPath);

                             Perform.SelectDropDown("//select[contains(@id,'ddSex_3')]", ExcelUtil.GetCellData(row, 77, Sheetname), Property_type.XPath);

                             Perform.SelectDropDown("//select[contains(@id,'ddMaritialStatus_3')]", ExcelUtil.GetCellData(row, 78, Sheetname), Property_type.XPath);
                             Perform.EnterText("//input[contains(@id,'txtDLNumber_3')]", ExcelUtil.GetCellData(row, 79, Sheetname), Property_type.XPath);
                             Perform.SelectDropDown(".//*[@id='cphMain_ctl_Master_Edit_ctlDriverList_Repeater1_ctlDriver_PPAControl_3_ddDLState_3']", ExcelUtil.GetCellData(row, 80, Sheetname), Property_type.XPath);
                             Perform.SelectDropDown(".//*[@id='cphMain_ctl_Master_Edit_ctlDriverList_Repeater1_ctlDriver_PPAControl_3_ddRelationToPolicyHolder_3']", ExcelUtil.GetCellData(row, 81, Sheetname), Property_type.XPath);
                             Perform.SelectDropDown(".//*[@id='cphMain_ctl_Master_Edit_ctlDriverList_Repeater1_ctlDriver_PPAControl_3_ddRatedOrExcludedDriver_3']", ExcelUtil.GetCellData(row, 82, Sheetname), Property_type.XPath);
                             Console.WriteLine("Driver4 details Entered");
                             Perform.ScreenShot(@"C:\Users\imnay\Documents\Visual Studio 2015\Projects\ClassLibrary1\ClassLibrary1\Output\DriverDetails.png");
                             Perform.Click(".//*[@id='cphMain_ctl_Master_Edit_ctlDriverList_btnSaveAndGotoVehicles']", Property_type.XPath);
                         }
                     }
                 }
             }*/

            //Perform.ScreenShot(@"C:\Users\imnay\Documents\Visual Studio 2015\Projects\ClassLibrary1\ClassLibrary1\Output\DriverDetails.png");

        }

        [Test]
        public void VehiclesPage(int row, String Sheetname, string savescreenshot)
        {
            DriverPage(row, "Driver", savescreenshot);
            Perform.test = Perform.report.StartTest("Vehicles Page");
            Perform.test.Log(LogStatus.Info, "Enter Vehicle Details");
            String nofdriver = ExcelUtil.GetCellData(row, 2, "TestCase");
            String nofvehicle = ExcelUtil.GetCellData(row, 3, "TestCase");


            PageUtility.AddVehicleDetails(row, nofdriver, nofvehicle, Sheetname);
            Perform.ScreenShot(savescreenshot + "VehicleDetails.png");
            Console.WriteLine("Vehicle details entered");
            Perform.test.Log(LogStatus.Info, " Vehicle Details Entered");
            Perform.test.Log(LogStatus.Info, "Proceed to Coverages");
            Perform.Click(".//*[@id='cphMain_ctl_Master_Edit_ctlVehicleList_btnSaveandGotoCoverages']", Property_type.XPath);
            Perform.IsElementPresent(".//*[@id='cphMain_ctl_Master_Edit_ctlCoverage_PPA_ctlCoverage_PPA_Vehicle_List_btnRate']");

            /* string result = ExcelUtil.GetCellData(row, 0, Sheetname);
             if (result == "")
             {
                 Perform.ScreenShot(@"C:\Users\imnay\Documents\Visual Studio 2015\Projects\ClassLibrary1\ClassLibrary1\Output\VehicleDetails.png");
                 Perform.Click(".//*[@id='cphMain_ctl_Master_Edit_ctlVehicleList_btnSaveandGotoCoverages']", Property_type.XPath);
             }
             else
             {
                 //Vehicles Page
                 Perform.Click(".//*[@id='cphMain_ctl_Master_Edit_ctlVehicleList_btnAddvehicle']", Property_type.XPath);
                 //Vehicle1 Details
                 Perform.EnterText(".//*[@id='cphMain_ctl_Master_Edit_ctlVehicleList_Repeater1_ctlVehicle_PPAControl_0_txtVinNumber_0']", ExcelUtil.GetCellData(row, 0, Sheetname), Property_type.XPath);
                 Perform.EnterText("//input[contains(@id,'txtYear_0')]", ExcelUtil.GetCellData(row, 1, Sheetname), Property_type.XPath);
                 Perform.EnterText(".//*[@id='cphMain_ctl_Master_Edit_ctlVehicleList_Repeater1_ctlVehicle_PPAControl_0_txtMake_0']", ExcelUtil.GetCellData(row, 2, Sheetname), Property_type.XPath);
                 Perform.EnterText(".//*[@id='cphMain_ctl_Master_Edit_ctlVehicleList_Repeater1_ctlVehicle_PPAControl_0_txtModel_0']", ExcelUtil.GetCellData(row, 3, Sheetname), Property_type.XPath);

                 Perform.Vehicletype(".//*[@id='cphMain_ctl_Master_Edit_ctlVehicleList_Repeater1_ctlVehicle_PPAControl_0_ddBodyType_0']", ExcelUtil.GetCellData(row, 6, Sheetname), ExcelUtil.GetCellData(row, 8, Sheetname), ExcelUtil.GetCellData(row, 9, Sheetname), ExcelUtil.GetCellData(row, 10, Sheetname));
                 Perform.SelectDropDown(".//*[@id='cphMain_ctl_Master_Edit_ctlVehicleList_Repeater1_ctlVehicle_PPAControl_0_ddPrincipalDriver_0']", ExcelUtil.GetCellData(row, 11, Sheetname), Property_type.XPath);
                 Console.WriteLine("Principal driver selected");

                 if (Property_Collection.driver.PageSource.Contains("Occasional Driver 1") && ExcelUtil.GetCellData(row, 12, Sheetname) != "")
                 {
                     Perform.SelectDropDown(".//*[@id='cphMain_ctl_Master_Edit_ctlVehicleList_Repeater1_ctlVehicle_PPAControl_0_ddOccDriver1_0']", ExcelUtil.GetCellData(row, 12, Sheetname), Property_type.XPath);
                     Console.WriteLine("Occasional driver 1 is selected");
                 }
                 else
                     Console.WriteLine("No Occassional Driver1");

                 if (Property_Collection.driver.PageSource.Contains("Occasional Driver 2") && ExcelUtil.GetCellData(row, 13, Sheetname) != "")
                 {
                     Perform.SelectDropDown(".//*[@id='cphMain_ctl_Master_Edit_ctlVehicleList_Repeater1_ctlVehicle_PPAControl_0_ddOccDriver2_0']", ExcelUtil.GetCellData(row, 13, Sheetname), Property_type.XPath);
                     Console.WriteLine("Occasional driver 2 selected");
                 }
                 else
                     Console.WriteLine("No Occassional driver2");

                 Perform.Click(".//*[@id='ui-id-28']", Property_type.XPath);
                 Perform.EnterText(".//*[@id='cphMain_ctl_Master_Edit_ctlVehicleList_Repeater1_ctlVehicle_PPAControl_0_txtGaragedStreetNum_0']", ExcelUtil.GetCellData(row, 17, Sheetname), Property_type.XPath);
                 Perform.EnterText(".//*[@id='cphMain_ctl_Master_Edit_ctlVehicleList_Repeater1_ctlVehicle_PPAControl_0_txtGaragedStreet_0']", ExcelUtil.GetCellData(row, 18, Sheetname), Property_type.XPath);
                 Perform.EnterText(".//*[@id='cphMain_ctl_Master_Edit_ctlVehicleList_Repeater1_ctlVehicle_PPAControl_0_txtGaragedApt_0']", ExcelUtil.GetCellData(row, 19, Sheetname), Property_type.XPath);
                 Perform.EnterText(".//*[@id='cphMain_ctl_Master_Edit_ctlVehicleList_Repeater1_ctlVehicle_PPAControl_0_txtGaragedCity_0']", ExcelUtil.GetCellData(row, 20, Sheetname), Property_type.XPath);
                 Perform.SelectDropDown(".//*[@id='cphMain_ctl_Master_Edit_ctlVehicleList_Repeater1_ctlVehicle_PPAControl_0_ddGaragedState_0']", ExcelUtil.GetCellData(row, 21, Sheetname), Property_type.XPath);
                 Perform.EnterText(".//*[@id='cphMain_ctl_Master_Edit_ctlVehicleList_Repeater1_ctlVehicle_PPAControl_0_txtGaragedZip_0']", ExcelUtil.GetCellData(row, 22, Sheetname), Property_type.XPath);
                 Perform.EnterText(".//*[@id='cphMain_ctl_Master_Edit_ctlVehicleList_Repeater1_ctlVehicle_PPAControl_0_txtGaragedCounty_0']", ExcelUtil.GetCellData(row, 23, Sheetname), Property_type.XPath);



                 Console.WriteLine("Vehicle1 details entered");

                 string result1 = ExcelUtil.GetCellData(row, 24, Sheetname);

                 if (result1 == "")
                 {
                     Perform.ScreenShot(@"C:\Users\imnay\Documents\Visual Studio 2015\Projects\ClassLibrary1\ClassLibrary1\Output\VehicleDetails.png");
                     Perform.waitTillElementToAppear(".//*[@id='cphMain_ctl_Master_Edit_ctlVehicleList_btnSaveandGotoCoverages']");
                     Perform.Click(".//*[@id='cphMain_ctl_Master_Edit_ctlVehicleList_btnSaveandGotoCoverages']", Property_type.XPath);
                 }
                 else
                 {
                     Perform.Click(".//*[@id='cphMain_ctl_Master_Edit_ctlVehicleList_btnAddvehicle']", Property_type.XPath);
                     Perform.EnterText(".//*[@id='cphMain_ctl_Master_Edit_ctlVehicleList_Repeater1_ctlVehicle_PPAControl_1_txtVinNumber_1']", ExcelUtil.GetCellData(row, 24, Sheetname), Property_type.XPath);
                     Perform.EnterText("//input[contains(@id,'txtYear_1')]", ExcelUtil.GetCellData(row, 25, Sheetname), Property_type.XPath);
                     Perform.EnterText(".//*[@id='cphMain_ctl_Master_Edit_ctlVehicleList_Repeater1_ctlVehicle_PPAControl_1_txtMake_1']", ExcelUtil.GetCellData(row, 26, Sheetname), Property_type.XPath);
                     Perform.EnterText(".//*[@id='cphMain_ctl_Master_Edit_ctlVehicleList_Repeater1_ctlVehicle_PPAControl_1_txtModel_1']", ExcelUtil.GetCellData(row, 27, Sheetname), Property_type.XPath);

                     Perform.Vehicletype(".//*[@id='cphMain_ctl_Master_Edit_ctlVehicleList_Repeater1_ctlVehicle_PPAControl_1_ddBodyType_1']", ExcelUtil.GetCellData(row, 30, Sheetname), ExcelUtil.GetCellData(row, 32, Sheetname), ExcelUtil.GetCellData(row, 33, Sheetname), ExcelUtil.GetCellData(row, 34, Sheetname));

                     Perform.SelectDropDown(".//*[@id='cphMain_ctl_Master_Edit_ctlVehicleList_Repeater1_ctlVehicle_PPAControl_1_ddPrincipalDriver_1']", ExcelUtil.GetCellData(row, 35, Sheetname), Property_type.XPath);
                     Console.WriteLine("Principal Driver selected");
                     if (Property_Collection.driver.PageSource.Contains("Occasional Driver 1") && ExcelUtil.GetCellData(row, 36, Sheetname) != "")
                     {

                         Perform.SelectDropDown(".//*[@id='cphMain_ctl_Master_Edit_ctlVehicleList_Repeater1_ctlVehicle_PPAControl_1_ddOccDriver1_1']", ExcelUtil.GetCellData(row, 36, Sheetname), Property_type.XPath);
                         Console.WriteLine("Occasional driver 1 is selected");
                     }

                     else
                         Console.WriteLine("No Occasional driver1");
                     if (Property_Collection.driver.PageSource.Contains("Occasional Driver 2") && ExcelUtil.GetCellData(row, 37, Sheetname) != "")
                     {


                         Perform.SelectDropDown(".//*[@id='cphMain_ctl_Master_Edit_ctlVehicleList_Repeater1_ctlVehicle_PPAControl_1_ddOccDriver2_1']", ExcelUtil.GetCellData(row, 37, Sheetname), Property_type.XPath);
                         Console.WriteLine("Occasional driver2 selected for 2nd vehicle");
                     }

                     else
                         Console.WriteLine("No occasional driver2");
                     Perform.Wait();
                     Console.WriteLine("Clicking on Garage information");
                     Perform.Click(".//*[@id='ui-id-42']/span[1]", Property_type.XPath);
                     //Console.WriteLine("Clicked on Garage information");
             Perform.waitTillElementToAppear(".//*[@id='cphMain_ctl_Master_Edit_ctlVehicleList_Repeater1_ctlVehicle_PPAControl_1_txtGaragedStreetNum_1']");
                     Perform.EnterText(".//*[@id='cphMain_ctl_Master_Edit_ctlVehicleList_Repeater1_ctlVehicle_PPAControl_1_txtGaragedStreetNum_1']", ExcelUtil.GetCellData(row, 41, Sheetname), Property_type.XPath);
                     Perform.EnterText(".//*[@id='cphMain_ctl_Master_Edit_ctlVehicleList_Repeater1_ctlVehicle_PPAControl_1_txtGaragedStreet_1']", ExcelUtil.GetCellData(row, 42, Sheetname), Property_type.XPath);
                     Perform.EnterText(".//*[@id='cphMain_ctl_Master_Edit_ctlVehicleList_Repeater1_ctlVehicle_PPAControl_1_txtGaragedApt_1']", ExcelUtil.GetCellData(row, 43, Sheetname), Property_type.XPath);
                     Perform.EnterText(".//*[@id='cphMain_ctl_Master_Edit_ctlVehicleList_Repeater1_ctlVehicle_PPAControl_1_txtGaragedCity_1']", ExcelUtil.GetCellData(row, 44, Sheetname), Property_type.XPath);
                     Perform.SelectDropDown(".//*[@id='cphMain_ctl_Master_Edit_ctlVehicleList_Repeater1_ctlVehicle_PPAControl_1_ddGaragedState_1']", ExcelUtil.GetCellData(row, 45, Sheetname), Property_type.XPath);
                     Perform.EnterText(".//*[@id='cphMain_ctl_Master_Edit_ctlVehicleList_Repeater1_ctlVehicle_PPAControl_1_txtGaragedZip_1']", ExcelUtil.GetCellData(row, 46, Sheetname), Property_type.XPath);
                     Perform.EnterText(".//*[@id='cphMain_ctl_Master_Edit_ctlVehicleList_Repeater1_ctlVehicle_PPAControl_1_txtGaragedCounty_1']", ExcelUtil.GetCellData(row, 47, Sheetname), Property_type.XPath);
                     Console.WriteLine("Vehicle2 details entered");



                     string result2 = ExcelUtil.GetCellData(row, 48, Sheetname);
                     if (result2 == "")
                     {
                         Perform.ScreenShot(@"C:\Users\imnay\Documents\Visual Studio 2015\Projects\ClassLibrary1\ClassLibrary1\Output\VehicleDetails.png");
                         Perform.Click(".//*[@id='cphMain_ctl_Master_Edit_ctlVehicleList_btnSaveandGotoCoverages']", Property_type.XPath);
                     }
                     else
                     {
                         Perform.Click(".//*[@id='cphMain_ctl_Master_Edit_ctlVehicleList_btnAddvehicle']", Property_type.XPath);
                         Perform.EnterText(".//*[@id='cphMain_ctl_Master_Edit_ctlVehicleList_Repeater1_ctlVehicle_PPAControl_2_txtVinNumber_2']", ExcelUtil.GetCellData(row, 48, Sheetname), Property_type.XPath);
                         Perform.EnterText("//input[contains(@id,'txtYear_2')]", ExcelUtil.GetCellData(row, 49, Sheetname), Property_type.XPath);
                         Perform.EnterText(".//*[@id='cphMain_ctl_Master_Edit_ctlVehicleList_Repeater1_ctlVehicle_PPAControl_2_txtMake_2']", ExcelUtil.GetCellData(row, 50, Sheetname), Property_type.XPath);
                         Perform.EnterText(".//*[@id='cphMain_ctl_Master_Edit_ctlVehicleList_Repeater1_ctlVehicle_PPAControl_2_txtModel_2']", ExcelUtil.GetCellData(row, 51, Sheetname), Property_type.XPath);


                         Perform.Vehicletype(".//*[@id='cphMain_ctl_Master_Edit_ctlVehicleList_Repeater1_ctlVehicle_PPAControl_2_ddBodyType_2']", ExcelUtil.GetCellData(row, 54, Sheetname), ExcelUtil.GetCellData(row, 56, Sheetname), ExcelUtil.GetCellData(row, 57, Sheetname), ExcelUtil.GetCellData(row, 58, Sheetname));
                         Perform.SelectDropDown(".//*[@id='cphMain_ctl_Master_Edit_ctlVehicleList_Repeater1_ctlVehicle_PPAControl_2_ddPrincipalDriver_2']", ExcelUtil.GetCellData(row, 59, Sheetname), Property_type.XPath);
                         if (Property_Collection.driver.PageSource.Contains("Occasional Driver 1") && ExcelUtil.GetCellData(row, 60, Sheetname) != "")
                         {
                             Perform.SelectDropDown(".//*[@id='cphMain_ctl_Master_Edit_ctlVehicleList_Repeater1_ctlVehicle_PPAControl_2_ddOccDriver1_2']", ExcelUtil.GetCellData(row, 60, Sheetname), Property_type.XPath);
                             Console.WriteLine("Occasional driver 1 is selected");
                             if (Property_Collection.driver.PageSource.Contains("Occasional Driver 2") && ExcelUtil.GetCellData(row, 61, Sheetname) != "")
                             {
                                 Perform.SelectDropDown(".//*[@id='cphMain_ctl_Master_Edit_ctlVehicleList_Repeater1_ctlVehicle_PPAControl_2_ddOccDriver2_2']", ExcelUtil.GetCellData(row, 61, Sheetname), Property_type.XPath);
                                 Console.WriteLine("Occasional driver2 selected for 3rd vehicle");
                             }
                         }

                         Perform.Click(".//*[@id='ui-id-56']", Property_type.XPath);
                         Perform.waitTillElementToAppear(".//*[@id='cphMain_ctl_Master_Edit_ctlVehicleList_Repeater1_ctlVehicle_PPAControl_2_txtGaragedStreetNum_2']");
                         Perform.EnterText(".//*[@id='cphMain_ctl_Master_Edit_ctlVehicleList_Repeater1_ctlVehicle_PPAControl_2_txtGaragedStreetNum_2']", ExcelUtil.GetCellData(row, 65, Sheetname), Property_type.XPath);
                         Perform.EnterText(".//*[@id='cphMain_ctl_Master_Edit_ctlVehicleList_Repeater1_ctlVehicle_PPAControl_2_txtGaragedStreet_2']", ExcelUtil.GetCellData(row, 66, Sheetname), Property_type.XPath);
                         Perform.EnterText(".//*[@id='cphMain_ctl_Master_Edit_ctlVehicleList_Repeater1_ctlVehicle_PPAControl_2_txtGaragedStreet_2']", ExcelUtil.GetCellData(row, 67, Sheetname), Property_type.XPath);
                         Perform.EnterText(".//*[@id='cphMain_ctl_Master_Edit_ctlVehicleList_Repeater1_ctlVehicle_PPAControl_2_txtGaragedCity_2']", ExcelUtil.GetCellData(row, 68, Sheetname), Property_type.XPath);
                         Perform.SelectDropDown(".//*[@id='cphMain_ctl_Master_Edit_ctlVehicleList_Repeater1_ctlVehicle_PPAControl_2_ddGaragedState_2']", ExcelUtil.GetCellData(row, 69, Sheetname), Property_type.XPath);
                         Perform.EnterText(".//*[@id='cphMain_ctl_Master_Edit_ctlVehicleList_Repeater1_ctlVehicle_PPAControl_2_txtGaragedZip_2']", ExcelUtil.GetCellData(row, 70, Sheetname), Property_type.XPath);
                         Perform.EnterText(".//*[@id='cphMain_ctl_Master_Edit_ctlVehicleList_Repeater1_ctlVehicle_PPAControl_2_txtGaragedCounty_2']", ExcelUtil.GetCellData(row, 71, Sheetname), Property_type.XPath);
                         Console.WriteLine("Vehicle3 details entered");
                         Perform.ScreenShot(@"C:\Users\imnay\Documents\Visual Studio 2015\Projects\ClassLibrary1\ClassLibrary1\Output\VehicleDetails.png");


                     }
                 }
             }

            // Perform.ScreenShot(@"C:\Users\imnay\Documents\Visual Studio 2015\Projects\ClassLibrary1\ClassLibrary1\Output\VehicleDetails.png");
             Console.WriteLine("Vehicle details entered");*/
        }

        [Test]
        public void Coverage(int row, String Sheetname, string savescreenshot)
        {
            String nofvehicle = ExcelUtil.GetCellData(row, 3, "TestCase");
            VehiclesPage(row, "Vehicle", savescreenshot);
            Perform.test = Perform.report.StartTest("Coverage Page");
            Perform.test.Log(LogStatus.Info, "Enter Vehicle Coverage Details");
            PageUtility.Coveragevalue(row, nofvehicle, Sheetname);

            //Coverage Page

            /* Perform.SelectDropDown(".//*[@id='cphMain_ctl_Master_Edit_ctlCoverage_PPA_ddLiabType']", ExcelUtil.GetCellData(row, 0, Sheetname), Property_type.XPath);
             Perform.SelectDropDown(".//*[@id='cphMain_ctl_Master_Edit_ctlCoverage_PPA_ddBodilyInjury']", ExcelUtil.GetCellData(row, 1, Sheetname), Property_type.XPath);

             Perform.SelectDropDown(".//*[@id='cphMain_ctl_Master_Edit_ctlCoverage_PPA_ddPropertyDamage']", ExcelUtil.GetCellData(row, 2, Sheetname), Property_type.XPath);
             Perform.SelectDropDown(".//*[@id='cphMain_ctl_Master_Edit_ctlCoverage_PPA_ddmedicalPayments']", ExcelUtil.GetCellData(row, 3, Sheetname), Property_type.XPath);
             Perform.SelectDropDown(".//*[@id='cphMain_ctl_Master_Edit_ctlCoverage_PPA_ddUmUmiBi']", ExcelUtil.GetCellData(row, 4, Sheetname), Property_type.XPath);

             Perform.SelectDropDown(".//*[@id='cphMain_ctl_Master_Edit_ctlCoverage_PPA_ddUmPd']", ExcelUtil.GetCellData(row, 5, Sheetname), Property_type.XPath);
             Perform.SelectDropDown(".//*[@id='cphMain_ctl_Master_Edit_ctlCoverage_PPA_ddUmPdDeductible']", ExcelUtil.GetCellData(row, 6, Sheetname), Property_type.XPath);

             Perform.SelectDropDown(".//*[@id='ddPolicy0']", ExcelUtil.GetCellData(row, 11, Sheetname), Property_type.XPath);
             if (ExcelUtil.GetCellData(row, 11, Sheetname) == "FULL COVERAGE")
             {

                 Perform.SelectDropDown(".//*[@id='cphMain_ctl_Master_Edit_ctlCoverage_PPA_ctlCoverage_PPA_Vehicle_List_Repeater1_ctlCoverage_PPA_VehicleSpecific_0_ddComprehensive_0']", ExcelUtil.GetCellData(row, 12, Sheetname), Property_type.XPath);
                 Perform.SelectDropDown(".//*[@id='cphMain_ctl_Master_Edit_ctlCoverage_PPA_ctlCoverage_PPA_Vehicle_List_Repeater1_ctlCoverage_PPA_VehicleSpecific_0_ddCollision_0']", ExcelUtil.GetCellData(row, 13, Sheetname), Property_type.XPath);
                 Perform.SelectDropDown(".//*[@id='cphMain_ctl_Master_Edit_ctlCoverage_PPA_ctlCoverage_PPA_Vehicle_List_Repeater1_ctlCoverage_PPA_VehicleSpecific_0_ddTowing_0']", ExcelUtil.GetCellData(row, 14, Sheetname), Property_type.XPath);
                 Perform.SelectDropDown(".//*[@id='cphMain_ctl_Master_Edit_ctlCoverage_PPA_ctlCoverage_PPA_Vehicle_List_Repeater1_ctlCoverage_PPA_VehicleSpecific_0_ddTransportation_0']", ExcelUtil.GetCellData(row, 15, Sheetname), Property_type.XPath);
                 Perform.SelectDropDown(".//*[@id='cphMain_ctl_Master_Edit_ctlCoverage_PPA_ctlCoverage_PPA_Vehicle_List_Repeater1_ctlCoverage_PPA_VehicleSpecific_0_ddRadio_0']", ExcelUtil.GetCellData(row, 16, Sheetname), Property_type.XPath);

                 Perform.SelectDropDown(".//*[@id='cphMain_ctl_Master_Edit_ctlCoverage_PPA_ctlCoverage_PPA_Vehicle_List_Repeater1_ctlCoverage_PPA_VehicleSpecific_0_ddAudioVisual_0']", ExcelUtil.GetCellData(row, 17, Sheetname), Property_type.XPath);
                 Perform.SelectDropDown(".//*[@id='cphMain_ctl_Master_Edit_ctlCoverage_PPA_ctlCoverage_PPA_Vehicle_List_Repeater1_ctlCoverage_PPA_VehicleSpecific_0_ddMedia_0']", ExcelUtil.GetCellData(row, 18, Sheetname), Property_type.XPath);

             }

             string result1 = ExcelUtil.GetCellData(row, 20, Sheetname);
             //Check if 2nd vehicle details are present
             if (result1 == "")
             {
                 Perform.Click(".//*[@id='cphMain_ctl_Master_Edit_ctlCoverage_PPA_ctlCoverage_PPA_Vehicle_List_btnRate']", Property_type.XPath);
             }
             else
             {

                 Perform.Click(".//*[@id='cphMain_ctl_Master_Edit_ctlCoverage_PPA_ctlCoverage_PPA_Vehicle_List_Repeater1_ctlCoverage_PPA_VehicleSpecific_1_lblAccordHeader_1']", Property_type.XPath);

                 Perform.SelectDropDown(".//*[@id='ddPolicy1']", ExcelUtil.GetCellData(row, 20, Sheetname), Property_type.XPath);

                 if (ExcelUtil.GetCellData(row, 20, Sheetname) == "FULL COVERAGE")
                 {
                     Console.WriteLine("VEHICLE 2" + ExcelUtil.GetCellData(row, 20, Sheetname));
                     System.Threading.Thread.Sleep(1000);
                     Perform.SelectDropDown(".//*[@id='cphMain_ctl_Master_Edit_ctlCoverage_PPA_ctlCoverage_PPA_Vehicle_List_Repeater1_ctlCoverage_PPA_VehicleSpecific_1_ddComprehensive_1']", ExcelUtil.GetCellData(row, 21, Sheetname), Property_type.XPath);

                     Perform.SelectDropDown(".//*[@id='cphMain_ctl_Master_Edit_ctlCoverage_PPA_ctlCoverage_PPA_Vehicle_List_Repeater1_ctlCoverage_PPA_VehicleSpecific_1_ddCollision_1']", ExcelUtil.GetCellData(row, 22, Sheetname), Property_type.XPath);

                     Perform.SelectDropDown(".//*[@id='cphMain_ctl_Master_Edit_ctlCoverage_PPA_ctlCoverage_PPA_Vehicle_List_Repeater1_ctlCoverage_PPA_VehicleSpecific_1_ddTowing_1']", ExcelUtil.GetCellData(row, 23, Sheetname), Property_type.XPath);

                     //Perform.SelectDropDown(".//*[@id='cphMain_ctl_Master_Edit_ctlCoverage_PPA_ctlCoverage_PPA_Vehicle_List_Repeater1_ctlCoverage_PPA_VehicleSpecific_1_ddTransportation_1']", ExcelUtil.GetCellData(row, 23, Sheetname), Property_type.XPath);

                     Perform.EnterText(".//*[@id='cphMain_ctl_Master_Edit_ctlCoverage_PPA_ctlCoverage_PPA_Vehicle_List_Repeater1_ctlCoverage_PPA_VehicleSpecific_1_txtMotorEquip_1']", ExcelUtil.GetCellData(row, 25, Sheetname), Property_type.XPath);

                     //Perform.SelectDropDown(".//*[@id='cphMain_ctl_Master_Edit_ctlCoverage_PPA_ctlCoverage_PPA_Vehicle_List_Repeater1_ctlCoverage_PPA_VehicleSpecific_1_ddAudioVisual_1']", ExcelUtil.GetCellData(row, 25, Sheetname), Property_type.XPath);

                     Perform.SelectDropDown(".//*[@id='cphMain_ctl_Master_Edit_ctlCoverage_PPA_ctlCoverage_PPA_Vehicle_List_Repeater1_ctlCoverage_PPA_VehicleSpecific_1_ddMedia_1']", ExcelUtil.GetCellData(row, 27, Sheetname), Property_type.XPath);

                 }

                 string result2 = ExcelUtil.GetCellData(row, 26, Sheetname);
                 //Check if 3rd vehicle details are present
                 if (result2 == "")
                 {
                     Perform.Click(".//*[@id='cphMain_ctl_Master_Edit_ctlCoverage_PPA_ctlCoverage_PPA_Vehicle_List_btnRate']", Property_type.XPath);
                 }
                 else
                 {
                     Perform.Click(".//*[@id='cphMain_ctl_Master_Edit_ctlCoverage_PPA_ctlCoverage_PPA_Vehicle_List_Repeater1_ctlCoverage_PPA_VehicleSpecific_2_lblAccordHeader_2']", Property_type.XPath);
                     Perform.SelectDropDown(".//*[@id='ddPolicy2']", ExcelUtil.GetCellData(row, 27, Sheetname), Property_type.XPath);
                     if (ExcelUtil.GetCellData(row, 29, Sheetname) == "FULL COVERAGE")
                     {
                         Console.WriteLine("vEHICLE 3");
                         Perform.SelectDropDown(".//*[@id='cphMain_ctl_Master_Edit_ctlCoverage_PPA_ctlCoverage_PPA_Vehicle_List_Repeater1_ctlCoverage_PPA_VehicleSpecific_2_ddComprehensive_2']", ExcelUtil.GetCellData(row, 30, Sheetname), Property_type.XPath);
                         Perform.SelectDropDown(".//*[@id='cphMain_ctl_Master_Edit_ctlCoverage_PPA_ctlCoverage_PPA_Vehicle_List_Repeater1_ctlCoverage_PPA_VehicleSpecific_2_ddCollision_2']", ExcelUtil.GetCellData(row, 31, Sheetname), Property_type.XPath);
                         Perform.SelectDropDown(".//*[@id='cphMain_ctl_Master_Edit_ctlCoverage_PPA_ctlCoverage_PPA_Vehicle_List_Repeater1_ctlCoverage_PPA_VehicleSpecific_2_ddTowing_2']", ExcelUtil.GetCellData(row, 32, Sheetname), Property_type.XPath);
                         Perform.SelectDropDown(".//*[@id='cphMain_ctl_Master_Edit_ctlCoverage_PPA_ctlCoverage_PPA_Vehicle_List_Repeater1_ctlCoverage_PPA_VehicleSpecific_2_ddTransportation_2']", ExcelUtil.GetCellData(row, 33, Sheetname), Property_type.XPath);
                         Perform.SelectDropDown(".//*[@id='cphMain_ctl_Master_Edit_ctlCoverage_PPA_ctlCoverage_PPA_Vehicle_List_Repeater1_ctlCoverage_PPA_VehicleSpecific_2_ddRadio_2']", ExcelUtil.GetCellData(row, 34, Sheetname), Property_type.XPath);

                         Perform.SelectDropDown(".//*[@id='cphMain_ctl_Master_Edit_ctlCoverage_PPA_ctlCoverage_PPA_Vehicle_List_Repeater1_ctlCoverage_PPA_VehicleSpecific_2_ddAudioVisual_2']", ExcelUtil.GetCellData(row, 35, Sheetname), Property_type.XPath);
                         Perform.SelectDropDown(".//*[@id='cphMain_ctl_Master_Edit_ctlCoverage_PPA_ctlCoverage_PPA_Vehicle_List_Repeater1_ctlCoverage_PPA_VehicleSpecific_2_ddMedia_2']", ExcelUtil.GetCellData(row, 36, Sheetname), Property_type.XPath);
                         Perform.Click(".//*[@id='cphMain_ctl_Master_Edit_ctlCoverage_PPA_ctlCoverage_PPA_Vehicle_List_btnRate']", Property_type.XPath);
                     }
                     Perform.Click(".//*[@id='cphMain_ctl_Master_Edit_ctlCoverage_PPA_ctlCoverage_PPA_Vehicle_List_btnRate']", Property_type.XPath);
                 }

             }*/
            Console.WriteLine("Coverage Details Entered");
            Perform.test.Log(LogStatus.Info, " Vehicle Coverage Details Entered");
           
            Perform.Click(".//*[@id='cphMain_ctl_Master_Edit_ctlCoverage_PPA_ctlCoverage_PPA_Vehicle_List_btnRate']", Property_type.XPath);
            Perform.IsElementPresent(".//*[@id='cphMain_ctl_Master_Edit_ctlQsummary_PPA_ctlQuoteSummaryActions_btnContinueToApp']");
            Perform.test.Log(LogStatus.Info, "Vehicle Quote");
            Perform.ScreenShot(savescreenshot + "Coverage.png");
          
         
            Perform.waitTillElementToAppear(".//*[@id='cphMain_ctl_Master_Edit_ctlQsummary_PPA_ctlQuoteSummaryActions_btnContinueToApp']");

            Perform.ScreenShot(savescreenshot + "QuoteSummary.png");
            // Perform.waitTillElementToAppear(".//*[@id='cphMain_ctl_Master_Edit_ctlQsummary_PPA_ctlQuoteSummaryActions_btnContinueToApp']");
            Perform.Click(".//*[@id='cphMain_ctl_Master_Edit_ctlQsummary_PPA_ctlQuoteSummaryActions_btnContinueToApp']", Property_type.XPath);
            Perform.IsElementPresent(".//*[@id='cphMain_ctl_App_Master_Edit_ctlUWQuestions_btnGoToApp']");
        }
        [Test]
        public void Underwriting(int row, String Sheetname, string savescreenshot)
        {
            Coverage(row, "Coverage", savescreenshot);
            Perform.test = Perform.report.StartTest("Underwriting Questions Page");
            Perform.test.Log(LogStatus.Info, " Click NO to all questions");
            //Underwriting Questions
            Perform.click_on_webElements("//input[contains(@id,'rbNo_')]", Property_Collection.driver);
            Perform.ScreenShot(savescreenshot + "Underwriting.png");
            Console.WriteLine("Answered Underwriting Questions");
            Perform.test.Log(LogStatus.Info, "Underwriring Questions answered");
            Perform.Click(".//*[@id='cphMain_ctl_App_Master_Edit_ctlUWQuestions_btnGoToApp']", Property_type.XPath);
            Perform.IsElementPresent(".//*[@id='btnShowEffectiveDate']");
        }

        [Test]
        public void BillingInfoPage(int row, String Sheetname, string savescreenshot)
        {
            Underwriting(row, " ", savescreenshot);
            Perform.test = Perform.report.StartTest("Billing Info Page");
            Perform.test.Log(LogStatus.Info, "Select Billing Info");
            Perform.SelectDropDown(".//*[@id='cphMain_ctl_App_Master_Edit_ctl_App_Section_ctl_Billing_Info_PPA_ddMethod']", ExcelUtil.GetCellData(row, 0, Sheetname), Property_type.XPath);
            Perform.SelectDropDown(".//*[@id='cphMain_ctl_App_Master_Edit_ctl_App_Section_ctl_Billing_Info_PPA_ddPayPlan']", ExcelUtil.GetCellData(row, 1, Sheetname), Property_type.XPath);
            Perform.SelectDropDown(".//*[@id='cphMain_ctl_App_Master_Edit_ctl_App_Section_ctl_Billing_Info_PPA_ddBillTo']", ExcelUtil.GetCellData(row, 2, Sheetname), Property_type.XPath);
            Console.WriteLine("Pay Plan details entered");
          
            Perform.Click(".//*[@id='btnShowEffectiveDate']", Property_type.XPath);
          
            Console.WriteLine("Rate Quote Button Clicked");
            //Select Effective Date
            Perform.SelectDropDown(".//*[@id='ui-datepicker-div']/div[1]/div/select[1]", ExcelUtil.GetCellData(row, 3, Sheetname), Property_type.XPath);

            Perform.SelectDropDown(".//*[@id='ui-datepicker-div']/div[1]/div/select[2]", ExcelUtil.GetCellData(row, 4, Sheetname), Property_type.XPath);

            Perform.Click(ExcelUtil.GetCellData(row, 5, Sheetname), Property_type.LinkText);
            System.Threading.Thread.Sleep(500);
            Property_Collection.driver.FindElement(By.Id("btnEffectiveDateDone")).SendKeys(Keys.Enter);
            Perform.Wait();
            Console.WriteLine("Effective Date Entered");
            Perform.test.Log(LogStatus.Info, "Pay Plans and Efefctive Date Entered");
        }
        [Test]
        public void FinalizePage()
        {
            //  DateTime time = DateTime.Now;
            // string dateToday = "_date_" + time.ToString("yyyy-MM-dd") + "_time_" + time.ToString("HH-mm-ss");


            int sheetrownum = ExcelUtil.getRowCount("TestCase");
            // String drivercount = ExcelUtil.GetCellData(3,2,"TestCase");


            for (int i = 3; i < sheetrownum; i++)
            {
                System.IO.Directory.CreateDirectory(path + ExcelUtil.GetCellData(i, 0, "TestCase"));
                string savescreenshot = path + ExcelUtil.GetCellData(i, 0, "TestCase") + "\\";
                BillingInfoPage(i, "BillingInfo", savescreenshot);
                Perform.test = Perform.report.StartTest("Policy Issue Page");
                Property_Collection.driver.Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(30);
                Perform.Click(".//*[@id='cphMain_ctl_App_Master_Edit_ctlQsummary_PPA_ctlQuoteSummaryActions_btnContinueToApp']", Property_type.XPath);
                Perform.Click(".//*[@id='cphMain_ctl_App_Master_Edit_ctlQsummary_PPA_ctlQuoteSummaryActions_lnkFinalize']", Property_type.XPath);
                Console.WriteLine("Application Finalized");
                Perform.test.Log(LogStatus.Info, "Application Finalized");
                Perform.ScreenShot(savescreenshot + "Policy" + (i - 2) + "_" + DateTime.Now.ToString("yyyyMMddhhmmss") + ".png");
                Perform.test.Log(LogStatus.Pass, "Policy is issued");
             
                Perform.waitTillElementToAppear(".//*[@id='CrumbsLogoutLink']");
                Perform.Click(".//*[@id='CrumbsLogoutLink']", Property_type.XPath);
                Perform.test.Log(LogStatus.Info, "Logged out");
                Console.WriteLine("Policy" + (i - 2) + " issued");

            }


            /*catch (Exception e)
            {
                Console.WriteLine(e);
                Perform.ScreenShot(@"C:\Users\imnay\Documents\Visual Studio 2015\Projects\ClassLibrary1\ClassLibrary1\Output\Failurepage.png");
            }*/
        }





        [TearDown]
        public void Cleanup()
        {

            if (TestContext.CurrentContext.Result.Outcome != ResultState.Success)

                Perform.ScreenShot(path + "Failure.png");
            Property_Collection.driver.Close();
            Perform.report.EndTest(Perform.test);
            Perform.report.Flush();
        }
    }
}





















































   