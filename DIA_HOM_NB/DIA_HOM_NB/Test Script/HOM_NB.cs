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

namespace DIA_HOM_NB
{
    public class HOM_NB
    {
        String url;
        String path = @"C:\Users\imnay\Documents\Visual Studio 2015\Projects\DIA_HOM_NB\DIA_HOM_NB\Output\";
        String ReportPath= @"C:\Users\imnay\Documents\Visual Studio 2015\Projects\DIA_HOM_NB\DIA_HOM_NB\Report\";
        [SetUp]
        public void Initialize()
        {


            Perform.Browser("chrome");

           // url = "http://ifmoldrelease/DiamondWeb/controlloader.aspx?p=Headquarters";
            url = "http://ifmdiapatch/DiamondWeb/controlloader.aspx?p=Headquarters";
            ExcelUtil.setExcelFile(@"C:\Users\imnay\Documents\Visual Studio 2015\Projects\DIA_HOM_NB\DIA_HOM_NB\Excel\HOM_DIA_NB.xlsx");
            Perform.report = new ExtentReports(ReportPath + "Report.html", CultureInfo.GetCultureInfo("es-ES"), true, DisplayOrder.NewestFirst);
            Perform.report.LoadConfig(ReportPath + "extent-config.xml");

        }
        [Test]
        public void HomePage()
        {
            Perform.test = Perform.report.StartTest("Home Page");
            Perform.driver.Navigate().GoToUrl(url);
            Console.WriteLine("Browser Opened");
            Perform.Click(".//*[@id='PoliciesMenu']");
            Perform.Click(".//*[@id='PoliciesSubMenu_1']/tbody/tr/td[2]/div/a");
            Perform.Click(".//*[@id='NewPolicySubMenu_0']/tbody/tr/td[2]/a");
            Console.WriteLine("New client is clicked");
            Perform.test.Log(LogStatus.Info, "New Client is clicked from Policies Menu");
        }
        [Test]
        public void PolicyholderPage(int row, string sheetname, string savescreenshot)
        {
            HomePage();
            Perform.test = Perform.report.StartTest("Policyholder Page");
            Perform.test.Log(LogStatus.Info, "Enter Policyholder Info");
            //Policyholder1 details
            Perform.EnterText(".//*[@id='P_L_ClientSubmissionWithAddressAndPOBox_Client1InsName_First']", ExcelUtil.GetCellData(row, 0, sheetname));
            Perform.EnterText(".//*[@id='P_L_ClientSubmissionWithAddressAndPOBox_Client1InsName_Middle']", ExcelUtil.GetCellData(row, 1, sheetname));
            Perform.EnterText(".//*[@id='P_L_ClientSubmissionWithAddressAndPOBox_Client1InsName_Last']", ExcelUtil.GetCellData(row, 2, sheetname));
            Perform.EnterText(".//*[@id='P_L_ClientSubmissionWithAddressAndPOBox_Client1InsName_PersonalTaxNumber_PersonalTaxNumber']", ExcelUtil.GetCellData(row, 3, sheetname));
            Console.WriteLine("SSN is entered");
            Perform.Click(".//*[@id='P_L_ClientSubmissionWithAddressAndPOBox_Client1InsName_Sex_D_I']");

            string gender = ExcelUtil.GetCellData(row, 4, sheetname);

            if (gender == "M")
            {
                Console.WriteLine(gender);
                Perform.Click(".//*[@id='P_L_ClientSubmissionWithAddressAndPOBox_Client1InsName_Sex_D_DDD_L_LBI1T1']");
            }
            if (gender == "F")
            {
                Console.WriteLine(gender);
                Perform.Click(".//*[@id='P_L_ClientSubmissionWithAddressAndPOBox_Client1InsName_Sex_D_DDD_L_LBI2T1']");
            }
            Console.WriteLine("Gender is given");

            string maritalstatus = ExcelUtil.GetCellData(row, 5, sheetname);

            Perform.Click(".//*[@id='P_L_ClientSubmissionWithAddressAndPOBox_Client1InsName_MaritalStatus_D_I']");
            if (maritalstatus == "Divorced")
            {
                Perform.Click(".//*[@id='P_L_ClientSubmissionWithAddressAndPOBox_Client1InsName_MaritalStatus_D_DDD_L_LBI1T1']");
            }
            if (maritalstatus == "Married")
            {
                Perform.Click(".//*[@id='P_L_ClientSubmissionWithAddressAndPOBox_Client1InsName_MaritalStatus_D_DDD_L_LBI2T1']");
            }
            if (maritalstatus == "Single")
            {
                Perform.Click(".//*[@id='P_L_ClientSubmissionWithAddressAndPOBox_Client1InsName_MaritalStatus_D_DDD_L_LBI3T1']");
            }
            if (maritalstatus == "Widowed")
            {
                Perform.Click(".//*[@id='P_L_ClientSubmissionWithAddressAndPOBox_Client1InsName_MaritalStatus_D_DDD_L_LBI4T1']");
            }




            System.Threading.Thread.Sleep(500);
            Perform.EnterText(".//*[@id='P_L_ClientSubmissionWithAddressAndPOBox_Client1InsName_LicenseNumber_LicenseNumber']", ExcelUtil.GetCellData(row, 6, sheetname));
            if (ExcelUtil.GetCellData(row, 7, sheetname) != "")
            {
                Perform.SelectTextDropDown(".//*[@id='P_L_ClientSubmissionWithAddressAndPOBox_Client1InsName_LicenseState_D_I']", ExcelUtil.GetCellData(row, 7, sheetname));
            }
            //Perform.Click(".//*[@id='P_L_ClientSubmissionWithAddressAndPOBox_Client1InsName_BirthDate_BirthDate_B-1']");
            // Perform.Click(".//*[@id='P_L_ClientSubmissionWithAddressAndPOBox_Client1InsName_BirthDate_BirthDate_DDD_C_T']");
            // PageUtility.Calendar(row,10, sheetname);


            Perform.EnterTextFocus(".//*[@id='P_L_ClientSubmissionWithAddressAndPOBox_Client1InsName_BirthDate_BirthDate_I']", ExcelUtil.GetCellData(row, 8, sheetname));

            //Address
            if (ExcelUtil.GetCellData(row, 10, sheetname) != "")
            {
                Perform.waitTillElementToAppear(".//*[@id='P_L_ClientSubmissionWithAddressAndPOBox_ClientAddress_HouseNumber']");
                Perform.EnterText(".//*[@id='P_L_ClientSubmissionWithAddressAndPOBox_ClientAddress_HouseNumber']", ExcelUtil.GetCellData(row, 10, sheetname));
                Perform.EnterText(".//*[@id='P_L_ClientSubmissionWithAddressAndPOBox_ClientAddress_StreetName']", ExcelUtil.GetCellData(row, 11, sheetname));
                Perform.EnterText(".//*[@id='P_L_ClientSubmissionWithAddressAndPOBox_ClientAddress_ApartmentNumber']", ExcelUtil.GetCellData(row, 12, sheetname));
                Perform.EnterText(".//*[@id='P_L_ClientSubmissionWithAddressAndPOBox_ClientAddress_OtherInfo']", ExcelUtil.GetCellData(row, 13, sheetname));
                Perform.EnterText(".//*[@id='P_L_ClientSubmissionWithAddressAndPOBox_ClientAddress_City']", ExcelUtil.GetCellData(row, 14, sheetname));
                Perform.SelectTextDropDown(".//*[@id='P_L_ClientSubmissionWithAddressAndPOBox_ClientAddress_AddressState_D_I']", ExcelUtil.GetCellData(row, 15, sheetname));

                Perform.EnterText(".//*[@id='P_L_ClientSubmissionWithAddressAndPOBox_ClientAddress_ZipCode_mtxtMain']", ExcelUtil.GetCellData(row, 16, sheetname));
                Perform.EnterText(".//*[@id='P_L_ClientSubmissionWithAddressAndPOBox_ClientAddress_County']", ExcelUtil.GetCellData(row, 17, sheetname));
                Perform.EnterText(".//*[@id='P_L_ClientSubmissionWithAddressAndPOBox_ClientAddress_Township']", ExcelUtil.GetCellData(row, 18, sheetname));
            }
            else
            {
                Perform.Click(".//*[@id='P_L_ClientSubmissionWithAddressAndPOBox_AddressTypeRadioButtonList_1']");
                Perform.waitTillElementToAppear(".//*[@id='P_L_ClientSubmissionWithAddressAndPOBox_ClientAddressPOBox_PostOfficeBox']");
                Perform.EnterText(".//*[@id='P_L_ClientSubmissionWithAddressAndPOBox_ClientAddressPOBox_PostOfficeBox']", ExcelUtil.GetCellData(row, 19, sheetname));
                Perform.EnterText(".//*[@id='P_L_ClientSubmissionWithAddressAndPOBox_ClientAddressPOBox_OtherInfo']", ExcelUtil.GetCellData(row, 20, sheetname));
                Perform.EnterText(".//*[@id='P_L_ClientSubmissionWithAddressAndPOBox_ClientAddress_City']", ExcelUtil.GetCellData(row, 21, sheetname));
                Perform.SelectTextDropDown(".//*[@id='P_L_ClientSubmissionWithAddressAndPOBox_ClientAddress_AddressState_D_I']", ExcelUtil.GetCellData(row, 22, sheetname));

                Perform.EnterText(".//*[@id='P_L_ClientSubmissionWithAddressAndPOBox_ClientAddress_ZipCode_mtxtMain']", ExcelUtil.GetCellData(row, 23, sheetname));
                Perform.EnterText(".//*[@id='P_L_ClientSubmissionWithAddressAndPOBox_ClientAddress_County']", ExcelUtil.GetCellData(row, 24, sheetname));
                Perform.EnterText(".//*[@id='P_L_ClientSubmissionWithAddressAndPOBox_ClientAddress_Township']", ExcelUtil.GetCellData(row, 25, sheetname));

            }
            Console.WriteLine("Policyholder1 details entered");
            Perform.test.Log(LogStatus.Info, "Policyholder1 Details Entered");
            if (ExcelUtil.GetCellData(row, 26, sheetname) != "")
            {
                Perform.Click(".//*[@id='ClientAccordion']/h3[2]/a");



                Perform.EnterText(".//*[@id='P_L_ClientSubmissionWithAddressAndPOBox_Client2InsName_First']", ExcelUtil.GetCellData(row, 26, sheetname));
                Perform.EnterText(".//*[@id='P_L_ClientSubmissionWithAddressAndPOBox_Client2InsName_Middle']", ExcelUtil.GetCellData(row, 27, sheetname));
                Perform.EnterText(".//*[@id='P_L_ClientSubmissionWithAddressAndPOBox_Client2InsName_Last']", ExcelUtil.GetCellData(row, 28, sheetname));
                System.Threading.Thread.Sleep(500);
                Perform.EnterText(".//*[@id='P_L_ClientSubmissionWithAddressAndPOBox_Client2InsName_PersonalTaxNumber_PersonalTaxNumber']", ExcelUtil.GetCellData(row, 29, sheetname));
                Perform.Click(".//*[@id='P_L_ClientSubmissionWithAddressAndPOBox_Client2InsName_Sex_D_I']");

                string gender1 = ExcelUtil.GetCellData(row, 30, sheetname);

                if (gender1 == "M")
                {
                    Console.WriteLine(gender);
                    Perform.Click(".//*[@id='P_L_ClientSubmissionWithAddressAndPOBox_Client2InsName_Sex_D_DDD_L_LBI1T1']");
                }
                if (gender1 == "F")
                {
                    Console.WriteLine(gender);
                    Perform.Click(".//*[@id='P_L_ClientSubmissionWithAddressAndPOBox_Client2InsName_Sex_D_DDD_L_LBI2T1']");
                }
                Console.WriteLine("Gender is given");

                string maritalstatus1 = ExcelUtil.GetCellData(row, 31, sheetname);

                Perform.Click(".//*[@id='P_L_ClientSubmissionWithAddressAndPOBox_Client2InsName_MaritalStatus_D_I']");
                if (maritalstatus == "Divorced")
                {
                    Perform.Click(".//*[@id='P_L_ClientSubmissionWithAddressAndPOBox_Client2InsName_MaritalStatus_D_DDD_L_LBI1T1']");
                }
                if (maritalstatus1 == "Married")
                {
                    Perform.Click(".//*[@id='P_L_ClientSubmissionWithAddressAndPOBox_Client2InsName_MaritalStatus_D_DDD_L_LBI2T1']");
                }
                if (maritalstatus1 == "Single")
                {
                    Perform.Click(".//*[@id='P_L_ClientSubmissionWithAddressAndPOBox_Client2InsName_MaritalStatus_D_DDD_L_LBI3T1']");
                }
                if (maritalstatus1 == "Widowed")
                {
                    Perform.Click(".//*[@id='P_L_ClientSubmissionWithAddressAndPOBox_Client2InsName_MaritalStatus_D_DDD_L_LBI4T1']");
                }

                Perform.EnterText(".//*[@id='P_L_ClientSubmissionWithAddressAndPOBox_Client2InsName_LicenseNumber_LicenseNumber']", ExcelUtil.GetCellData(row, 32, sheetname));
                if (ExcelUtil.GetCellData(row, 33, sheetname) != "")
                {
                    Perform.SelectTextDropDown(".//*[@id='P_L_ClientSubmissionWithAddressAndPOBox_Client2InsName_LicenseState_D_I']", ExcelUtil.GetCellData(row, 33, sheetname));
                }
                Perform.EnterTextFocus(".//*[@id='P_L_ClientSubmissionWithAddressAndPOBox_Client2InsName_BirthDate_BirthDate_I']", ExcelUtil.GetCellData(row, 34, sheetname));

                Console.WriteLine("Policyholder2 details entered");
                Perform.test.Log(LogStatus.Info, "Policyholder2 Details Entered");
            }

            Perform.ScreenShot(savescreenshot + "Policyholder.png");
            Perform.Click(".//*[@id='ContinueInsImageButtonMiddle']/a");
            System.Threading.Thread.Sleep(2000);
        }
        [Test]
        public void TransactionPage(int row, string sheetname, string savescreenshot)
        {
            PolicyholderPage(row, "New_Client_Entry", savescreenshot);
            Perform.test = Perform.report.StartTest("Transaction Information Page");
            //Enter transaction info

            Perform.waitTillElementToAppear(".//*[@id='P_L_Application_C0044_VersionEffectiveDateInsDateTime_VersionEffectiveDateInsDateTime']");
            Perform.EnterTextFocus(".//*[@id='P_L_Application_C0044_VersionEffectiveDateInsDateTime_VersionEffectiveDateInsDateTime']", ExcelUtil.GetCellData(row, 0, sheetname));
            Console.WriteLine("Date entered");
            Perform.test.Log(LogStatus.Info, "Effective Date Entered");
            Perform.EnterTextFocus(".//*[@id='P_L_Application_C0044_LobInsCombo_D_I']", ExcelUtil.GetCellData(row, 1, sheetname));
            System.Threading.Thread.Sleep(1000);
            Perform.driver.FindElement(By.XPath(".//*[@id='P_L_Application_C0044_LobInsCombo_D_I']")).SendKeys(Keys.Tab);
            System.Threading.Thread.Sleep(2000);
         
            Perform.EnterTextMove(".//*[@id='P_L_Application_C0044_AgencyInsCombo_D_I']", ExcelUtil.GetCellData(row, 2, sheetname));
            System.Threading.Thread.Sleep(1000);
            Perform.test.Log(LogStatus.Info, "Agency is selected");
            Perform.ClickPerform(".//*[@id='P_L_Application_C0044_AgencyInsCombo_D_DDD_L_LBI0T0']/em");
            Console.WriteLine("Agency Code Entered");
            System.Threading.Thread.Sleep(1000);
            Perform.test.Log(LogStatus.Info, "Transaction Information is given");
            Perform.ScreenShot(savescreenshot + "TransactionInfo.png");

            //Perform.ClickPerform("//a[contains(text(),'Next')]");
            Perform.driver.FindElement(By.XPath(".//*[@id='SubmitToolStripButtonTwoMiddle']/a")).SendKeys(Keys.PageDown);
            System.Threading.Thread.Sleep(100);
            Perform.Click(".//*[@id='SubmitToolStripButtonTwoMiddle']/a");
            System.Threading.Thread.Sleep(500);
        }
        [Test]
        public void PolicyLevelInfoPage(int row, string sheetname, string savescreenshot)
        {
            TransactionPage(row, "Transaction_Information", savescreenshot);
            Perform.test = Perform.report.StartTest("Policy Level Information Page");
            System.Threading.Thread.Sleep(1000);
            Perform.Click("//a[contains(text(),'Policy Level Information')]");
            Perform.test.Log(LogStatus.Info, "Enter Policy Level Info");
            Perform.EnterTextFocus(".//*[@id='P_L_V_v39w9_t16_c0w0_t0_OccurrenceLiabilityInsCombo_D_I']",ExcelUtil.GetCellData(row, 0, sheetname));
            Perform.EnterTextFocus(".//*[@id='P_L_V_v39w9_t16_c0w0_t0_MedPayInsCombo_D_I']",ExcelUtil.GetCellData(row, 1, sheetname));
         
          
            Perform.Click(".//*[@id='SaveTieringInfoToolstripButtonMiddle']/a");
            System.Threading.Thread.Sleep(500);
            Perform.test.Log(LogStatus.Info, "Policy level Info is given");
            Perform.ScreenShot(savescreenshot + "PolicyLevelInfo.png");
            System.Threading.Thread.Sleep(500);

        }
        [Test]
        public void LocationsPage(int row, string sheetname, string savescreenshot)
        {
            PolicyLevelInfoPage(row, "Policy_Level_Information", savescreenshot);
            Perform.test = Perform.report.StartTest("Location Page");
            Perform.Click("//a[contains(text(),'Locations')]");
            Perform.Click(".//*[@id='AddLocationToolStripButtonMiddle']/a");
            Perform.waitTillElementToAppear(".//*[@id='P_L_V_v39w9_t17_c0w0_PC_t0i0_c0w0_t1_YearBuiltInsNumeric_YearBuiltInsNumeric']");
            Perform.test.Log(LogStatus.Info, "Add New Location");
            //Change Property Address
            if (ExcelUtil.GetCellData(row, 0, sheetname) != "")
            {
                Perform.Click(".//*[@id='P_L_V_v39w9_t17_c0w0_PC_t0i0_c0w0_t0_ChangeAddressLinkButton']");
                Perform.waitTillElementToAppear(".//*[@id='P_L_V_v39w9_t17_c0w0_PC_t0i0_c0w0_t0_InsNameAddressControl_HouseNumber']");
                Perform.EnterText(".//*[@id='P_L_V_v39w9_t17_c0w0_PC_t0i0_c0w0_t0_InsNameAddressControl_HouseNumber']", ExcelUtil.GetCellData(row, 1, sheetname));
                Perform.EnterText(".//*[@id='P_L_V_v39w9_t17_c0w0_PC_t0i0_c0w0_t0_InsNameAddressControl_StreetName']", ExcelUtil.GetCellData(row, 2, sheetname));
                Perform.EnterText(".//*[@id='P_L_V_v39w9_t17_c0w0_PC_t0i0_c0w0_t0_InsNameAddressControl_ApartmentNumber']", ExcelUtil.GetCellData(row, 3, sheetname));
                Perform.EnterText(".//*[@id='P_L_V_v39w9_t17_c0w0_PC_t0i0_c0w0_t0_InsNameAddressControl_PostOfficeBox']", ExcelUtil.GetCellData(row, 4, sheetname));
                Perform.EnterText(".//*[@id='P_L_V_v39w9_t17_c0w0_PC_t0i0_c0w0_t0_InsNameAddressControl_OtherInfo']", ExcelUtil.GetCellData(row, 5, sheetname));
                Perform.EnterText(".//*[@id='P_L_V_v39w9_t17_c0w0_PC_t0i0_c0w0_t0_InsNameAddressControl_City']", ExcelUtil.GetCellData(row, 6, sheetname));
                Perform.SelectDropDown(".//*[@id='P_L_V_v39w9_t17_c0w0_PC_t0i0_c0w0_t0_InsNameAddressControl_AddressState_D_I']", ExcelUtil.GetCellData(row, 7, sheetname));
                Perform.EnterText(".//*[@id='P_L_V_v39w9_t17_c0w0_PC_t0i0_c0w0_t0_InsNameAddressControl_ZipCode_mtxtMain']", ExcelUtil.GetCellData(row, 8, sheetname));
                Perform.EnterText(".//*[@id='P_L_V_v39w9_t17_c0w0_PC_t0i0_c0w0_t0_InsNameAddressControl_County']", ExcelUtil.GetCellData(row, 9, sheetname));
                Perform.Click(".//*[@id='SaveAddressInsImageButtonMiddle']/a");
                System.Threading.Thread.Sleep(500);
            }
            //General Information
            if (ExcelUtil.GetCellData(row, 10, sheetname) == "NO")
            {
                Perform.Click(".//*[@id='P_L_V_v39w9_t17_c0w0_PC_t0i0_c0w0_t1_PrimaryResidenceInsCheckBox']");
            }
            Perform.EnterTextFocus(".//*[@id='P_L_V_v39w9_t17_c0w0_PC_t0i0_c0w0_t1_FormTypeInsCombo_D_I']", ExcelUtil.GetCellData(row, 11, sheetname));
            System.Threading.Thread.Sleep(1000);
            //Perform.driver.FindElement(By.XPath(".//*[@id='P_L_V_v39w9_t17_c0w0_PC_t0i0_c0w0_t1_FormTypeInsCombo_D_I']")).SendKeys(Keys.Tab);
           // Perform.driver.SwitchTo().Alert().Accept();
            Perform.EnterText(".//*[@id='P_L_V_v39w9_t17_c0w0_PC_t0i0_c0w0_t1_YearBuiltInsNumeric_YearBuiltInsNumeric']", ExcelUtil.GetCellData(row, 12, sheetname));
            Perform.EnterText(".//*[@id='P_L_V_v39w9_t17_c0w0_PC_t0i0_c0w0_t1_SquareFeetInsNumeric_SquareFeetInsNumeric']", ExcelUtil.GetCellData(row, 13, sheetname));
            Perform.EnterTextFocus(".//*[@id='P_L_V_v39w9_t17_c0w0_PC_t0i0_c0w0_t1_NumberOfFamiliesInsCombo_D_I']", ExcelUtil.GetCellData(row, 14, sheetname));
            Perform.EnterTextFocus(".//*[@id='P_L_V_v39w9_t17_c0w0_PC_t0i0_c0w0_t1_StructureTypeInsCombo_D_I']", ExcelUtil.GetCellData(row, 15, sheetname));
           System.Threading.Thread.Sleep(1000);
            Perform.driver.FindElement(By.XPath(".//*[@id='P_L_V_v39w9_t17_c0w0_PC_t0i0_c0w0_t1_OccupancyCodeInsCombo_D_I']")).SendKeys(ExcelUtil.GetCellData(row, 16, sheetname));
           // Perform.waitTillElementToAppear(".//*[@id='P_L_V_v39w9_t17_c0w0_PC_t0i0_c0w0_t1_OccupancyCodeInsCombo_D_I']");
            //Perform.Click(".//*[@id='P_L_V_v39w9_t17_c0w0_PC_t0i0_c0w0_t1_OccupancyCodeInsCombo_D_I']");
            //Perform.EnterTextFocus(".//*[@id='P_L_V_v39w9_t17_c0w0_PC_t0i0_c0w0_t1_OccupancyCodeInsCombo_D_I']", ExcelUtil.GetCellData(row, 16, sheetname));
            System.Threading.Thread.Sleep(500);
            Perform.EnterTextFocus(".//*[@id='P_L_V_v39w9_t17_c0w0_PC_t0i0_c0w0_t1_ConstructionTypeInsCombo_D_I']", ExcelUtil.GetCellData(row, 17, sheetname));
            Perform.EnterTextFocus(".//*[@id='P_L_V_v39w9_t17_c0w0_PC_t0i0_c0w0_t1_FoundationTypeInsCombo_D_I']", ExcelUtil.GetCellData(row, 18, sheetname));
            if (ExcelUtil.GetCellData(row, 19, sheetname) != "")
            {
                Perform.EnterTextFocus(".//*[@id='P_L_V_v39w9_t17_c0w0_PC_t0i0_c0w0_t1_CondoRentedTypeInsCombo_D_I']", ExcelUtil.GetCellData(row, 19, sheetname));
            }
            if (ExcelUtil.GetCellData(row, 20, sheetname) != "")
            {
                Perform.EnterTextFocus(".//*[@id='P_L_V_v39w9_t17_c0w0_PC_t0i0_c0w0_t1_NumberOfSolidFuelBurningUnitsInsNumeric_NumberOfSolidFuelBurningUnitsInsNumeric']", ExcelUtil.GetCellData(row, 20, sheetname));
            }
            if (ExcelUtil.GetCellData(row, 21, sheetname) != "")
            {
                Perform.EnterTextFocus(".//*[@id='P_L_V_v39w9_t17_c0w0_PC_t0i0_c0w0_t1_FireDepartmentDistanceInsCombo_D_I']", ExcelUtil.GetCellData(row, 21, sheetname));
            }
            Perform.EnterTextFocus(".//*[@id='P_L_V_v39w9_t17_c0w0_PC_t0i0_c0w0_t1_FireHydrantDistanceInsCombo_D_I']", ExcelUtil.GetCellData(row, 22, sheetname));
            //Coverage Info
            Perform.EnterTextFocus(".//*[@id='P_L_V_v39w9_t17_c0w0_PC_t0i0_c0w0_t2_DeductibleInsCombo_D_I']", ExcelUtil.GetCellData(row, 23, sheetname));
            Perform.EnterTextFocus(".//*[@id='P_L_V_v39w9_t17_c0w0_PC_t0i0_c0w0_t2_WindHailDedInsCombo_D_I']", ExcelUtil.GetCellData(row, 24, sheetname));
            //Limit
            System.Threading.Thread.Sleep(500);
            Perform.EnterTextFocus(".//*[@id='P_L_V_v39w9_t17_c0w0_PC_t0i0_c0w0_t2_CovALimitInsNumeric_CovALimitInsNumeric_I']", ExcelUtil.GetCellData(row, 25, sheetname));
            if (ExcelUtil.GetCellData(row, 26, sheetname) != "")
            {
                Perform.EnterTextFocus(".//*[@id='P_L_V_v39w9_t17_c0w0_PC_t0i0_c0w0_t2_CovBLimitInsNumeric_CovBLimitInsNumeric_I']", ExcelUtil.GetCellData(row, 26, sheetname));
            }
            if (ExcelUtil.GetCellData(row, 27, sheetname) != "")
            {
                Perform.EnterTextFocus(".//*[@id='P_L_V_v39w9_t17_c0w0_PC_t0i0_c0w0_t2_CovCLimitInsNumeric_CovCLimitInsNumeric_I']", ExcelUtil.GetCellData(row, 27, sheetname));
            }
            if (ExcelUtil.GetCellData(row, 28, sheetname) != "")
            {
                Perform.EnterTextFocus(".//*[@id='P_L_V_v39w9_t17_c0w0_PC_t0i0_c0w0_t2_CovDLimitInsNumeric_CovDLimitInsNumeric_I']", ExcelUtil.GetCellData(row, 28, sheetname));
            }
            //Change in Limit
            if (ExcelUtil.GetCellData(row, 30, sheetname) != "")
            {
                Perform.EnterTextFocus(".//*[@id='P_L_V_v39w9_t17_c0w0_PC_t0i0_c0w0_t2_CovBIncLimitInsNumeric_CovBIncLimitInsNumeric_I']", ExcelUtil.GetCellData(row, 30, sheetname));
            }
            if (ExcelUtil.GetCellData(row, 31, sheetname) != "")
            {
                Perform.EnterTextFocus(".//*[@id='P_L_V_v39w9_t17_c0w0_PC_t0i0_c0w0_t2_CovCIncLimitInsNumeric_CovCIncLimitInsNumeric_I']", ExcelUtil.GetCellData(row, 31, sheetname));
            }
            if (ExcelUtil.GetCellData(row, 32, sheetname) != "")
            {
                Perform.EnterTextFocus(".//*[@id='P_L_V_v39w9_t17_c0w0_PC_t0i0_c0w0_t2_CovDIncLimitInsNumeric_CovDIncLimitInsNumeric_I']", ExcelUtil.GetCellData(row, 32, sheetname));
            }
            Console.WriteLine("Location:Coverage Info is entered");
            Perform.test.Log(LogStatus.Info, "Location General Info Given");
            System.Threading.Thread.Sleep(1000);
            //Updates
            if (ExcelUtil.GetCellData(row, 0, "Updates") != "" && ExcelUtil.GetCellData(row, 0, "Updates") == "YES")
            {
                PageUtility.Updates(row, sheetname);
                Console.WriteLine("Updates Info is entered");
                Perform.test.Log(LogStatus.Info, "Updates Info Given");
            }
           
            System.Threading.Thread.Sleep(500);
            //Additional Interest
            if (ExcelUtil.GetCellData(row, 0, "Additional Interest") != "" && ExcelUtil.GetCellData(row, 0, "Additional Interest") == "YES")
            {
                Perform.Click(".//*[@id='P_L_V_v39w9_t17_c0w0_PC_AT3T']/span");
                Perform.Click(".//*[@id='AddAdditionalInterestImageButtonMiddle']/a");
                System.Threading.Thread.Sleep(1000);
                Perform.EnterTextFocus(".//*[@id='P_L_V_v39w9_t17_c0w0_PC_t3i0_c0w0_t0i0_TypeInsCombo_D_I']", ExcelUtil.GetCellData(row, 1, "Additional Interest"));
                Perform.Click(".//*[@id='LookupToolStripButtonMiddle']/a");
                Perform.EnterText(".//*[@id='P_L_V_v39w9_t17_c0w0_PC_t3i0_c0w0_t1_NameTextBox']", ExcelUtil.GetCellData(row, 2, "Additional Interest"));
                Perform.EnterText(".//*[@id='P_L_V_v39w9_t17_c0w0_PC_t3i0_c0w0_t1_CityTextBox']", ExcelUtil.GetCellData(row, 3, "Additional Interest"));
                Perform.EnterTextFocus(".//*[@id='P_L_V_v39w9_t17_c0w0_PC_t3i0_c0w0_t1_InsState_D_I']", ExcelUtil.GetCellData(row, 4, "Additional Interest"));
                Perform.EnterText(".//*[@id='P_L_V_v39w9_t17_c0w0_PC_t3i0_c0w0_t1_ZipCode_mtxtMain']", ExcelUtil.GetCellData(row, 5, "Additional Interest"));
                Perform.Click(".//*[@id='FindLookupToolBarMiddle']/a");
                System.Threading.Thread.Sleep(1000);
                Perform.Click(".//*[@id='P_L_V_v39w9_t17_c0w0_PC_t3i0_c0w0_t1_AdditionalInterestLookupInsDataGridView_cell0_0_SelectASPxButton_0Img']");
                Perform.Click(".//*[@id='SaveToolStripButtonMiddle']/a");
                Perform.test.Log(LogStatus.Info, "Additional Interest Info Given");
            }
            System.Threading.Thread.Sleep(500);
           
            //Credits & Surcharges
            //Credits
            if (ExcelUtil.GetCellData(row, 0, "Credits_Surcharges") != "" && ExcelUtil.GetCellData(row, 0, "Credits_Surcharges") == "YES")
            {
                Perform.Click("//a[@id='P_L_V_v39w9_t17_c0w0_PC_T4T']/span");
                if (ExcelUtil.GetCellData(row, 1, "Credits_Surcharges") != "" && ExcelUtil.GetCellData(row, 1, "Credits_Surcharges") == "YES")
                {
                    Perform.waitTillElementToAppear(".//*[@id='P_L_V_v39w9_t17_c0w0_PC_t4i0_modifier_id_1_1_8_1_S_D']");
                    Perform.Click(".//*[@id='P_L_V_v39w9_t17_c0w0_PC_t4i0_modifier_id_1_1_8_1_S_D']");
                }
                if (ExcelUtil.GetCellData(row, 2, "Credits_Surcharges") != "" && ExcelUtil.GetCellData(row, 2, "Credits_Surcharges") == "YES")
                {
                    Perform.Click(".//*[@id='P_L_V_v39w9_t17_c0w0_PC_t4i0_modifier_id_3_1_8_3_S_D']");
                }
                if (ExcelUtil.GetCellData(row, 3, "Credits_Surcharges") != "" && ExcelUtil.GetCellData(row, 3, "Credits_Surcharges") == "YES")
                {
                    Perform.Click(".//*[@id='P_L_V_v39w9_t17_c0w0_PC_t4i0_modifier_id_5_1_8_5_S_D']");
                }
                if (ExcelUtil.GetCellData(row, 4, "Credits_Surcharges") != "" && ExcelUtil.GetCellData(row, 4, "Credits_Surcharges") == "YES")
                {
                    Perform.Click(".//*[@id='P_L_V_v39w9_t17_c0w0_PC_t4i0_modifier_id_36_1_8_36_S_D']");
                }
                //Burglar alarm
                if (ExcelUtil.GetCellData(row, 5, "Credits_Surcharges") != "" && ExcelUtil.GetCellData(row, 5, "Credits_Surcharges") == "YES")
                {
                    Perform.Click(".//*[@id='P_L_V_v39w9_t17_c0w0_PC_t4i0_modifier_id_10_1_8_26_S_D']");
                }
                if (ExcelUtil.GetCellData(row, 6, "Credits_Surcharges") != "" && ExcelUtil.GetCellData(row, 6, "Credits_Surcharges") == "YES")
                {
                    Perform.Click(".//*[@id='P_L_V_v39w9_t17_c0w0_PC_t4i0_modifier_id_9_1_8_26_S_D']");
                }
                //Fire/Smoke Alarm
                if (ExcelUtil.GetCellData(row, 7, "Credits_Surcharges") != "" && ExcelUtil.GetCellData(row, 7, "Credits_Surcharges") == "YES")
                {
                    Perform.Click(".//*[@id='P_L_V_v39w9_t17_c0w0_PC_t4i0_modifier_id_10_1_8_27_S_D']");
                }
                if (ExcelUtil.GetCellData(row, 8, "Credits_Surcharges") != "" && ExcelUtil.GetCellData(row, 8, "Credits_Surcharges") == "YES")
                {
                    Perform.Click(".//*[@id='P_L_V_v39w9_t17_c0w0_PC_t4i0_modifier_id_9_1_8_27_S_D']");
                }
                if (ExcelUtil.GetCellData(row, 9, "Credits_Surcharges") != "" && ExcelUtil.GetCellData(row, 9, "Credits_Surcharges") == "YES")
                {
                    Perform.Click(".//*[@id='P_L_V_v39w9_t17_c0w0_PC_t4i0_modifier_id_53_1_8_27_S_D']");
                }
                //Sprinkler System
                if (ExcelUtil.GetCellData(row, 10, "Credits_Surcharges") != "" && ExcelUtil.GetCellData(row, 10, "Credits_Surcharges") == "YES")
                {
                    Perform.Click(".//*[@id='P_L_V_v39w9_t17_c0w0_PC_t4i0_modifier_id_39_1_8_28_S_D']");
                }
                if (ExcelUtil.GetCellData(row, 11, "Credits_Surcharges") != "" && ExcelUtil.GetCellData(row, 11, "Credits_Surcharges") == "YES")
                {
                    Perform.Click(".//*[@id='P_L_V_v39w9_t17_c0w0_PC_t4i0_modifier_id_40_1_8_28_S_D']");
                }
                //Surcharges/Fee
                if (ExcelUtil.GetCellData(row, 12, "Credits_Surcharges") != "" && ExcelUtil.GetCellData(row, 12, "Credits_Surcharges") == "YES")
                {
                    Perform.Click(".//*[@id='P_L_V_v39w9_t17_c0w0_PC_t4i0_modifier_id_37_2_8_37_S_D']");
                }
                if (ExcelUtil.GetCellData(row, 13, "Credits_Surcharges") != "" && ExcelUtil.GetCellData(row, 13, "Credits_Surcharges") == "YES")
                {
                    Perform.Click(".//*[@id='P_L_V_v39w9_t17_c0w0_PC_t4i0_modifier_id_17_2_8_17_S_D']");
                }
                if (ExcelUtil.GetCellData(row, 14, "Credits_Surcharges") != "" && ExcelUtil.GetCellData(row, 14, "Credits_Surcharges") == "YES")
                {
                    Perform.Click(".//*[@id='P_L_V_v39w9_t17_c0w0_PC_t4i0_modifier_id_65_2_8_65_S_D']");
                }
                Console.WriteLine("Credits/Surcharge Info is given");
                Perform.test.Log(LogStatus.Info, "Credits/Surcharge Info Given");
            }
            Perform.Click(".//*[@id='SaveToolStripButtonMiddle']/a");
            System.Threading.Thread.Sleep(500);
            Perform.waitTillElementToAppear("//a[contains(text(),'Optional Coverages')]");
            //Optional Coverages
            Perform.Click("//a[contains(text(),'Optional Coverages')]");
             Console.WriteLine("Optional Coverage is clicked");
             if (Perform.driver.FindElement(By.XPath(".//*[@id='P_L_V_ValidationPopUp_MyASPxPopupControl_OKInsValidationButton_CD']/span")).Displayed)
                 {
                     Perform.Click(".//*[@id='P_L_V_ValidationPopUp_MyASPxPopupControl_OKInsValidationButton_CD']/span");
                 }
             System.Threading.Thread.Sleep(1000);
             //Section 1
             if (ExcelUtil.GetCellData(row, 0, "Optional_Coverages_Section_I") != "" && ExcelUtil.GetCellData(row, 0, "Optional_Coverages_Section_I") == "YES")
             {
              
                 string nofcoverage = ExcelUtil.GetCellData(row,1, "Optional_Coverages_Section_I");
                Console.WriteLine(nofcoverage);
                int cov = Int32.Parse(nofcoverage);
                 int col = 2;
                 for (int i = 0; i <cov; i++)
                 {
                     Perform.waitTillElementToAppear(".//*[@id='AddToolStripButtonMiddle']/a");
                     Perform.Click(".//*[@id='AddToolStripButtonMiddle']/a");
                     Console.WriteLine("Add New Coverage is clciked");
                     Perform.EnterTextFocus(".//*[@id='P_L_V_v39w9_t18_c0w0_PC_t0_c0w0_PC_t0_InsCoverageControl_CoverageInsCombo_D_I']", ExcelUtil.GetCellData(row, col, "Optional_Coverages_Section_I"));
                     if (ExcelUtil.GetCellData(row, col+1, "Optional_Coverages_Section_I") != "")
                     {
                         Perform.EnterText(".//*[@id='P_L_V_v39w9_t18_c0w0_PC_t0_c0w0_PC_t0_InsCoverageControl_CoverageControlASPxCallbackPanel_A_12111_877_A_12111_877_MainLimitLimit_A_12111_877_MainLimitLimit_I']", ExcelUtil.GetCellData(row, col + 1, "Optional_Coverages_Section_I"));
                     }
                     if (ExcelUtil.GetCellData(row, col+2, "Optional_Coverages_Section_I") != "")
                     {
                         Perform.EnterText(".//*[@id='P_L_V_v39w9_t18_c0w0_PC_t0_c0w0_PC_t0_DescriptionInsTextBox']", ExcelUtil.GetCellData(row, col + 2, "Optional_Coverages_Section_I"));
                     }
                     Perform.Click(".//*[@id='SaveToolStripButtonMiddle']/a");
                     col = col + 3;
                 }
                 System.Threading.Thread.Sleep(500);
                 Perform.ScreenShot(savescreenshot + "SectionI.png");
             }
             //Section 2
             if (ExcelUtil.GetCellData(row, 0, "Optional_Coverages_Section_II") != "" && ExcelUtil.GetCellData(row, 0, "Optional_Coverages_Section_II") == "YES")
             {
                 Perform.Click(".//*[@id='P_L_V_v39w9_t18_c0w0_PC_AT1T']/span");

                 int col = 2;
                 string nofcoverage2 = ExcelUtil.GetCellData(row, 1, " Optional_Coverages_Section_II");
                 for (int i = 0; i < Int32.Parse(nofcoverage2); i++)
                 {
                     Perform.Click(".//*[@id='AddToolStripButtonMiddle']/a");
                     Perform.EnterTextFocus(".//*[@id='P_L_V_v39w9_t18_c0w0_PC_t1_c0w0_PC_t0_InsCoverageControl_CoverageInsCombo_D_I']", ExcelUtil.GetCellData(row, col, "Optional_Coverages_Section_II"));
                     if (ExcelUtil.GetCellData(row, col + 1, "Optional_Coverages_Section_II") != "")
                     {
                         Perform.EnterText(".//*[@id='P_L_V_v39w9_t18_c0w0_PC_t1_c0w0_PC_t0_NumberOfPersonsReceivingCareInsNumeric_NumberOfPersonsReceivingCareInsNumeric']", ExcelUtil.GetCellData(row, col + 1, "Optional_Coverages_Section_II"));
                     }
                     if (ExcelUtil.GetCellData(row, col + 2, "Optional_Coverages_Section_II") != "")
                     {
                         Perform.EnterText(".//*[@id='P_L_V_v39w9_t18_c0w0_PC_t1_c0w0_PC_t0_NumberOfFamiliesInsNumeric_NumberOfFamiliesInsNumeric']", ExcelUtil.GetCellData(row, col + 2, "Optional_Coverages_Section_II"));
                     }
                     if (ExcelUtil.GetCellData(row, col + 3, "Optional_Coverages_Section_II") != "")
                     {
                         Perform.EnterText(".//*[@id='P_L_V_v39w9_t18_c0w0_PC_t1_c0w0_PC_t0_NumberOfFullTimeEmployeesInsNumeric_NumberOfFullTimeEmployeesInsNumeric']", ExcelUtil.GetCellData(row, col + 3, "Optional_Coverages_Section_II"));
                     }
                     if (ExcelUtil.GetCellData(row, col + 4, "Optional_Coverages_Section_II") != "")
                     {
                         Perform.EnterText(".//*[@id='P_L_V_v39w9_t18_c0w0_PC_t1_c0w0_PC_t0_NumberOfPartTimeEmployees41DaysInsNumeric_NumberOfPartTimeEmployees41DaysInsNumeric']", ExcelUtil.GetCellData(row, col + 4, "Optional_Coverages_Section_II"));
                     }
                     if (ExcelUtil.GetCellData(row, col + 5, "Optional_Coverages_Section_II") != "")
                     {
                         Perform.EnterText(".//*[@id='P_L_V_v39w9_t18_c0w0_PC_t1_c0w0_PC_t0_NumberOfPartTimeEmployees40DaysInsNumeric_NumberOfPartTimeEmployees40DaysInsNumeric']", ExcelUtil.GetCellData(row, col + 5, "Optional_Coverages_Section_II"));
                     }
                     if (ExcelUtil.GetCellData(row, col + 6, "Optional_Coverages_Section_II") != "")
                     {
                         Perform.EnterText(".//*[@id='P_L_V_v39w9_t18_c0w0_PC_t1_c0w0_PC_t0_NumberOfLivestockInsNumeric_NumberOfLivestockInsNumeric']", ExcelUtil.GetCellData(row, col + 6, "Optional_Coverages_Section_II"));
                     }
                     if (ExcelUtil.GetCellData(row, col + 7, "Optional_Coverages_Section_II") != "")
                     {
                         Perform.EnterText(".//*[@id='P_L_V_v39w9_t18_c0w0_PC_t1_c0w0_PC_t0_BusinessTypeInsTextBox']", ExcelUtil.GetCellData(row, col + 7, "Optional_Coverages_Section_II"));
                     }
                     Perform.Click(".//*[@id='SaveToolStripButtonMiddle']/a");
                     col = col + 8;
                 }
                 System.Threading.Thread.Sleep(500);
                 Perform.ScreenShot(savescreenshot + "SectionII.png");
             }
             //Section 1 & 2
             if (ExcelUtil.GetCellData(row, 0, "Optional_Coverages_Section_I_II") != "" && ExcelUtil.GetCellData(row, 0, "Optional_Coverages_Section_I_II") == "YES")
             {
                 Perform.Click(".//*[@id='P_L_V_v39w9_t18_c0w0_PC_AT2T']/span");
                 string nocoverage3 = ExcelUtil.GetCellData(row, 1, "Optional_Coverages_Section_I_II");
                 int col = 2;
                 for(int i=0;i<Int32.Parse(nocoverage3);i++)
                     {
                     Perform.Click(".//*[@id='AddToolStripButtonMiddle']/a");
                     Perform.EnterTextFocus(".//*[@id='P_L_V_v39w9_t18_c0w0_PC_t2_c0w0_PC_t0_InsCoverageControl_CoverageInsCombo_D_I']", ExcelUtil.GetCellData(row, col, "Optional_Coverages_Section_I_II"));
                     if (ExcelUtil.GetCellData(row, col + 1, "Optional_Coverages_Section_II") != "")
                     {
                         Perform.EnterText(".//*[@id='P_L_V_v39w9_t18_c0w0_PC_t2_c0w0_PC_t0_DescriptionInsTextBox']", ExcelUtil.GetCellData(row, col + 1, "Optional_Coverages_Section_I_II"));
                     }
                     if (ExcelUtil.GetCellData(row, col + 3, "Optional_Coverages_Section_I_II") != "")
                     {
                         Perform.EnterText(".//*[@id='P_L_V_v39w9_t18_c0w0_PC_t2_c0w0_PC_t0_NumberOfFamiliesInsNumeric_NumberOfFamiliesInsNumeric']", ExcelUtil.GetCellData(row, col + 2, "Optional_Coverages_Section_I_II"));
                     }
                     if (ExcelUtil.GetCellData(row, col + 4, "Optional_Coverages_Section_I_II") != "")
                     {
                         Perform.SelectTextDropDown(".//*[@id='P_L_V_v39w9_t18_c0w0_PC_t2_c0w0_PC_t0_EarthquakeZoneInsCombo_D_I']", ExcelUtil.GetCellData(row, col + 4, "Optional_Coverages_Section_I_II"));
                     }
                     col = col + 5;
                 }
                 System.Threading.Thread.Sleep(500);
                 Perform.ScreenShot(savescreenshot + "SectionI_II.png");
             }
            Perform.test.Log(LogStatus.Info, "Optional Coverage Info Given");
            //Inland Marine
            if (ExcelUtil.GetCellData(row,0,"Inland_Marine")!="")
            {
                System.Threading.Thread.Sleep(500);
                Perform.Click("//a[contains(text(),'Inland Marine')]");
                if(Perform.IsElementDisplayed(".//*[@id='P_L_V_ValidationPopUp_MyASPxPopupControl_OKInsValidationButton_CD']/span") == true)
                {
                    Perform.Click(".//*[@id='P_L_V_ValidationPopUp_MyASPxPopupControl_OKInsValidationButton_CD']/span");
                }

               /* if (Perform.driver.FindElement(By.XPath(".//*[@id='P_L_V_ValidationPopUp_MyASPxPopupControl_OKInsValidationButton_CD']/span")).Displayed)
                {
                    Perform.Click(".//*[@id='P_L_V_ValidationPopUp_MyASPxPopupControl_OKInsValidationButton_CD']/span");
                }*/
                PageUtility.InlandMarine(row, sheetname);
                System.Threading.Thread.Sleep(500);
                Perform.ScreenShot(savescreenshot + "InlandMarine.png");
            }
            Perform.test.Log(LogStatus.Info, "Inland Marine Info Given");
            //Watercraft
            if (ExcelUtil.GetCellData(row, 0, "R_V_Watercraft") != "")
            {
                Perform.Click(".//*[@id='P_L_V_DetailTreeViewn0Nodes']/table[7]/tbody/tr/td[3]/span/a[2]");
                if(Perform.IsElementDisplayed(".//*[@id='P_L_V_ValidationPopUp_MyASPxPopupControl_OKInsValidationButton_CD']/span")==true)
                {
                    Perform.Click(".//*[@id='P_L_V_ValidationPopUp_MyASPxPopupControl_OKInsValidationButton_CD']/span");
                }
                PageUtility.RVWatercraft(row, sheetname);
                System.Threading.Thread.Sleep(500);
                Perform.test.Log(LogStatus.Info, "Watercraft Info Given");
                Perform.ScreenShot(savescreenshot + "Watercraft.png");
            }
            System.Threading.Thread.Sleep(500);
        }
        [Test]
        public void AdditionalPolicyInfoPage(int row, string sheetname, string savescreenshot)
        {
            LocationsPage(row, "Locations", savescreenshot);
            Perform.test = Perform.report.StartTest("Additional Policy Info Page");
            System.Threading.Thread.Sleep(1000);
            // Perform.Click("//a[contains(text(),'Additonal Policy Info')]");
            Perform.Click(".//*[@id='P_L_V_DetailTreeViewn0Nodes']/table[8]/tbody/tr/td[3]/span/a[2]");
            if (Perform.IsElementDisplayed(".//*[@id='P_L_V_ValidationPopUp_MyASPxPopupControl_OKInsValidationButton_CD']/span")== true)
            {
                Perform.Click(".//*[@id='P_L_V_ValidationPopUp_MyASPxPopupControl_OKInsValidationButton_CD']/span");
            }
           /* if (Perform.driver.FindElement(By.XPath(".//*[@id='P_L_V_ValidationPopUp_MyASPxPopupControl_OKInsValidationButton_CD']/span")).Displayed)
            {
                Perform.Click(".//*[@id='P_L_V_ValidationPopUp_MyASPxPopupControl_OKInsValidationButton_CD']/span");
            }*/
            System.Threading.Thread.Sleep(500);
            Perform.SelectDropDown(".//*[@id='P_L_V_v39w9_t21_c0w0_NB_ITC0i0_t0_0_AvailableProductsListBox_0']", "CLUE");
            Perform.Click(".//*[@id='SelectProductButtonMiddle']/a");
            System.Threading.Thread.Sleep(500);
            Perform.SelectDropDown(".//*[@id='P_L_V_v39w9_t21_c0w0_NB_ITC0i0_t0_0_AvailableProductsListBox_0']", "Credit Report");
            Perform.Click(".//*[@id='SelectProductButtonMiddle']/a");
            System.Threading.Thread.Sleep(500);
            Perform.Click(".//span[contains(@id,'checkAllSubjectsCheckbox_0_S_D')]");
            Perform.Click(".//span[contains(@id,'checkAllRiskLocationsCheckbox_0_S_D')]");
            Perform.Click(".//*[@id='OrderToolstripButtonMiddle']/a");
            Perform.Click(".//*[@id='P_L_V_v39w9_t21_c0w0_NB_ITC0i0_t0_0_ChoicepointControlValidationList_0_MyASPxPopupControl_0_ContinueInsValidationButton_0_CD']/span");
            Console.WriteLine("Continue button is clicked");
            Perform.test.Log(LogStatus.Warning, "Continue button is clicked");
            Perform.waitTillElementToAppear(".//*[@id='P_L_V_ValidationPopUp_MyASPxPopupControl_OtherLabel']");
            //  Perform.driver.SwitchTo().Alert().Accept();
            Perform.Click(".//*[@id='P_L_V_ValidationPopUp_MyASPxPopupControl_OKInsValidationButton_CD']/span");
            System.Threading.Thread.Sleep(500);
            Perform.test.Log(LogStatus.Info, "OK button is clicked in Message Box");
            Perform.test.Log(LogStatus.Info, "Additional POlicy Info is Given");
            Perform.ScreenShot(savescreenshot + "AdditionalPolicyInfo.png");

        }
        [Test]
        public void BillingPage(int row, string sheetname, string savescreenshot)
        {
            AdditionalPolicyInfoPage(row, "Additional_Policy_Info", savescreenshot);
            Perform.test = Perform.report.StartTest("Billing Page");
            System.Threading.Thread.Sleep(1000);
            Perform.Click(".//*[@id='P_L_V_DetailTreeViewn0Nodes']/table[9]/tbody/tr/td[3]/span/a[2]");
            Perform.EnterTextFocus(".//*[@id='P_L_V_v39w9_t22_MethodComboBox_D_I']", ExcelUtil.GetCellData(row, 0, sheetname));
            System.Threading.Thread.Sleep(1000);
            Perform.EnterTextFocus(".//*[@id='P_L_V_v39w9_t22_PayPlanInsCombo_D_I']", ExcelUtil.GetCellData(row, 1, sheetname));
            System.Threading.Thread.Sleep(1000);
            Perform.EnterTextFocus(".//*[@id='P_L_V_v39w9_t22_BillToControl_BillToInsComboBox_D_I']", ExcelUtil.GetCellData(row, 2, sheetname));
            Console.WriteLine("Billing Info is given");
            System.Threading.Thread.Sleep(500);
            Perform.test.Log(LogStatus.Info, "Billing Details Entered");
            Perform.ScreenShot(savescreenshot + "BillingInfo.png");
        }
        [Test]
        public void UnderwritingPage(int row, string sheetname, string savescreenshot)
        {
            BillingPage(row, "Billing", savescreenshot);
            Perform.test = Perform.report.StartTest("Underwriting Page");
            System.Threading.Thread.Sleep(1000);
            Perform.Click(".//*[@id='P_L_V_DetailTreeViewn0Nodes']/table[10]/tbody/tr/td[3]/span/a[2]");
            Perform.test.Log(LogStatus.Info, "Answer Underwriting Questions");
            Perform.Click(".//*[@id='P_L_V_v39w9_t23_radiobutton_9446_104_1_No']");
            int j = 9299;
            for (int i = 0; i < 25; i++)
            {
                Perform.click_on_webElements("//input[contains(@id,'radiobutton_" + j + "_104_1_No')]", Perform.driver);
                j = j + 1;
            }
            System.Threading.Thread.Sleep(500);
            Perform.test.Log(LogStatus.Info, "Underwriting Questions Answered");
            Perform.ScreenShot(savescreenshot + "Underwriting.png");
        }
        [Test]
        public void ViewQuotePage()
        {
            int sheetrownum = ExcelUtil.getRowCount("Test_Case");
            try
            {
                for (int i = 4; i < sheetrownum; i++)
                {
                    System.IO.Directory.CreateDirectory(path + ExcelUtil.GetCellData(i, 0, "Test_Case"));
                    string savescreenshot = path + ExcelUtil.GetCellData(i, 0, "Test_Case") + "\\";
                    Console.WriteLine(savescreenshot);
                    UnderwritingPage(i, "Underwriting", savescreenshot);
                    Perform.test = Perform.report.StartTest("View Quote Page");
                    //View Quote
                    Perform.Click(".//*[@id='P_L_V_RateToolStripButton']");
                    Perform.Click("//div[@id='P_L_V_ValidationPopUp_MyASPxPopupControl_OKInsValidationButton_CD']/span");
                    /*if (Perform.driver.FindElement(By.XPath(".//*[@id='P_L_V_ValidationPopUp_MyASPxPopupControl_ContinueInsValidationButton_CD']/span")).Displayed)
                    {
                        Perform.Click(".//*[@id='P_L_V_ValidationPopUp_MyASPxPopupControl_ContinueInsValidationButton_CD']/span");
                    }*/
                    System.Threading.Thread.Sleep(1000);
                    Perform.Click(".//*[@id='P_L_V_DetailTreeViewn0Nodes']/table[11]/tbody/tr/td[3]/span/a[2]");
                    Perform.Click("//div[@id='P_L_V_ValidationPopUp_MyASPxPopupControl_OKInsValidationButton_CD']/span");
                    Perform.test.Log(LogStatus.Warning, "OK button Clicked");
                    System.Threading.Thread.Sleep(500);
                    Perform.ScreenShot(savescreenshot + "Policy" + (i - 3) + "_" + DateTime.Now.ToString("yyyyMMddhhmmss") + ".png");
                    Console.WriteLine("Policy" + (i - 3) + "is isssued");
                    Perform.test.Log(LogStatus.Pass, "Policy is Issued");
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("Exception" + e);
                Assert.True(false);
                Perform.test = Perform.report.StartTest("Error");
                Perform.test.Log(LogStatus.Fail, "Policy is not Issued");
            }

        }
        [TearDown]
        public void CleanUp()
        {
            if (TestContext.CurrentContext.Result.Outcome != ResultState.Success)
            {
                Perform.ScreenShot(path + "Failure.png");
                Perform.driver.Close();
                Perform.report.EndTest(Perform.test);
                Perform.report.Flush();
            }
        }
    }
}








