using NUnit.Framework;
using NUnit.Framework.Interfaces;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using RelevantCodes.ExtentReports;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DIA_PPA_NB
{
    public class PPA_NB
    {

        String url;
        String path = @"C:\Users\imnay\Documents\Visual Studio 2015\Projects\DIA_PPA_NB\DIA_PPA_NB\Output\";
        String ReportPath = @"C:\Users\imnay\Documents\Visual Studio 2015\Projects\DIA_PPA_NB\DIA_PPA_NB\Report\";

       [SetUp]
        public void Initialize()
        {


            Perform.Browser("IE");

            // url = "http://ifmoldrelease/DiamondWeb/controlloader.aspx?p=Headquarters";
            url = "http://ifmdiapatch/DiamondWeb/controlloader.aspx?p=Headquarters";
           // ExcelUtil.setExcelFile(@"C:\Users\imnay\Desktop\PPA_DIA_NB1.xlsx");
           ExcelUtil.setExcelFile(@"C:\Users\imnay\Documents\Visual Studio 2015\Projects\DIA_PPA_NB\DIA_PPA_NB\Excel\PPA_DIA_NB.xlsx");
            Perform.report = new ExtentReports(ReportPath + "Report.html", CultureInfo.GetCultureInfo("es-ES"), true, DisplayOrder.NewestFirst);
            //Perform.report.LoadConfig(ReportPath + "extent-config.xml");

        }
        [Test]
        public void HomePage()
        {
            Perform.test = Perform.report.StartTest("Home Page");

            Perform.driver.Navigate().GoToUrl(url);
            
            Console.WriteLine("Browser Opened");
            Perform.test.Log(LogStatus.Info, "Browser Opened");
            Perform.Click(".//*[@id='PoliciesMenu']");
            Perform.Click(".//*[@id='PoliciesSubMenu_1']/tbody/tr/td[2]/div/a");
            Perform.Click(".//*[@id='NewPolicySubMenu_0']/tbody/tr/td[2]/a");
            Console.WriteLine("New client is clicked");
            Perform.test.Log(LogStatus.Info, "New Client is clicked from Policies Menu");
            //Perform.PageContains("By entering SSN, a more accurate quote will be generated.");

        }
        [Test]
        public void PolicyholderPage(int row, string sheetname, string savescreenshot)
        {
            HomePage();
            Perform.test = Perform.report.StartTest("Policyholder Page");
            Perform.test.Log(LogStatus.Info, "Enter Policyholder Info");
            //Policyholder1 details

            Perform.SelectTextDropDown(".//*[@id='P_L_ClientSubmissionWithAddressAndPOBox_Client1InsName_Prefix_D_I']", ExcelUtil.GetCellData(row, 0, sheetname));

            Console.WriteLine("Prefix Entered");
            Perform.EnterText(".//*[@id='P_L_ClientSubmissionWithAddressAndPOBox_Client1InsName_First']", ExcelUtil.GetCellData(row, 1, sheetname));
            Perform.EnterText(".//*[@id='P_L_ClientSubmissionWithAddressAndPOBox_Client1InsName_Middle']", ExcelUtil.GetCellData(row, 2, sheetname));
            Perform.EnterText(".//*[@id='P_L_ClientSubmissionWithAddressAndPOBox_Client1InsName_Last']", ExcelUtil.GetCellData(row, 3, sheetname));
            Perform.SelectTextDropDown(".//*[@id='P_L_ClientSubmissionWithAddressAndPOBox_Client1InsName_Suffix_D_I']", ExcelUtil.GetCellData(row, 5, sheetname));
            Perform.EnterText(".//*[@id='P_L_ClientSubmissionWithAddressAndPOBox_Client1InsName_PersonalTaxNumber_PersonalTaxNumber']", ExcelUtil.GetCellData(row, 4, sheetname));
            Console.WriteLine("SSN is entered");
            Perform.Click(".//*[@id='P_L_ClientSubmissionWithAddressAndPOBox_Client1InsName_Sex_D_I']");

            string gender = ExcelUtil.GetCellData(row, 6, sheetname);

            if (gender == "Male")
            {
                Console.WriteLine(gender);
                Perform.Click(".//*[@id='P_L_ClientSubmissionWithAddressAndPOBox_Client1InsName_Sex_D_DDD_L_LBI1T1']");
            }
            if (gender == "Female")
            {
                Console.WriteLine(gender);
                Perform.Click(".//*[@id='P_L_ClientSubmissionWithAddressAndPOBox_Client1InsName_Sex_D_DDD_L_LBI2T1']");
            }
            Console.WriteLine("Gender is given");

            string maritalstatus = ExcelUtil.GetCellData(row, 7, sheetname);

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
            Perform.EnterText(".//*[@id='P_L_ClientSubmissionWithAddressAndPOBox_Client1InsName_LicenseNumber_LicenseNumber']", ExcelUtil.GetCellData(row, 8, sheetname));
            Perform.SelectTextDropDown(".//*[@id='P_L_ClientSubmissionWithAddressAndPOBox_Client1InsName_LicenseState_D_I']", ExcelUtil.GetCellData(row, 9, sheetname));

            //Perform.Click(".//*[@id='P_L_ClientSubmissionWithAddressAndPOBox_Client1InsName_BirthDate_BirthDate_B-1']");
            // Perform.Click(".//*[@id='P_L_ClientSubmissionWithAddressAndPOBox_Client1InsName_BirthDate_BirthDate_DDD_C_T']");
            // PageUtility.Calendar(row,10, sheetname);


            Perform.EnterTextFocus(".//*[@id='P_L_ClientSubmissionWithAddressAndPOBox_Client1InsName_BirthDate_BirthDate_I']", ExcelUtil.GetCellData(row, 10, sheetname));

            //Address
            Perform.waitTillElementToAppear(".//*[@id='P_L_ClientSubmissionWithAddressAndPOBox_ClientAddress_HouseNumber']");
            Perform.EnterText(".//*[@id='P_L_ClientSubmissionWithAddressAndPOBox_ClientAddress_HouseNumber']", ExcelUtil.GetCellData(row, 12, sheetname));
            Perform.EnterText(".//*[@id='P_L_ClientSubmissionWithAddressAndPOBox_ClientAddress_StreetName']", ExcelUtil.GetCellData(row, 13, sheetname));
            Perform.EnterText(".//*[@id='P_L_ClientSubmissionWithAddressAndPOBox_ClientAddress_ApartmentNumber']", ExcelUtil.GetCellData(row, 14, sheetname));
            Perform.EnterText(".//*[@id='P_L_ClientSubmissionWithAddressAndPOBox_ClientAddress_City']", ExcelUtil.GetCellData(row, 15, sheetname));
            Perform.SelectTextDropDown(".//*[@id='P_L_ClientSubmissionWithAddressAndPOBox_ClientAddress_AddressState_D_I']", ExcelUtil.GetCellData(row, 16, sheetname));

            Perform.EnterText(".//*[@id='P_L_ClientSubmissionWithAddressAndPOBox_ClientAddress_ZipCode_mtxtMain']", ExcelUtil.GetCellData(row, 17, sheetname));

            Console.WriteLine("Policyholder1 details entered");
            Perform.test.Log(LogStatus.Info, "Policyholder1 Details Entered");
            if (ExcelUtil.GetCellData(row, 18, sheetname) != "")
            {
                Perform.Click(".//*[@id='ClientAccordion']/h3[2]/a");
                Perform.SelectTextDropDown(".//*[@id='P_L_ClientSubmissionWithAddressAndPOBox_Client2InsName_Prefix_D_I']", ExcelUtil.GetCellData(row, 18, sheetname));

                Console.WriteLine("Prefix Entered");
                Perform.EnterText(".//*[@id='P_L_ClientSubmissionWithAddressAndPOBox_Client2InsName_First']", ExcelUtil.GetCellData(row, 19, sheetname));
                Perform.EnterText(".//*[@id='P_L_ClientSubmissionWithAddressAndPOBox_Client2InsName_Middle']", ExcelUtil.GetCellData(row, 20, sheetname));
                Perform.EnterText(".//*[@id='P_L_ClientSubmissionWithAddressAndPOBox_Client2InsName_Last']", ExcelUtil.GetCellData(row, 21, sheetname));
                Perform.SelectTextDropDown(".//*[@id='P_L_ClientSubmissionWithAddressAndPOBox_Client2InsName_Suffix_D_I']", ExcelUtil.GetCellData(row, 22, sheetname));
                System.Threading.Thread.Sleep(500);
                Perform.EnterText(".//*[@id='P_L_ClientSubmissionWithAddressAndPOBox_Client2InsName_PersonalTaxNumber_PersonalTaxNumber']", ExcelUtil.GetCellData(row, 23, sheetname));
                Perform.Click(".//*[@id='P_L_ClientSubmissionWithAddressAndPOBox_Client1InsName_Sex_D_I']");

                string gender1 = ExcelUtil.GetCellData(row, 24, sheetname);

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

                string maritalstatus1 = ExcelUtil.GetCellData(row, 25, sheetname);

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

                Perform.EnterText(".//*[@id='P_L_ClientSubmissionWithAddressAndPOBox_Client2InsName_LicenseNumber_LicenseNumber']", ExcelUtil.GetCellData(row, 26, sheetname));
                Perform.SelectTextDropDown(".//*[@id='P_L_ClientSubmissionWithAddressAndPOBox_Client2InsName_LicenseState_D_I']", ExcelUtil.GetCellData(row, 27, sheetname));
                Perform.EnterTextFocus(".//*[@id='P_L_ClientSubmissionWithAddressAndPOBox_Client2InsName_BirthDate_BirthDate_B-1']", ExcelUtil.GetCellData(row, 28, sheetname));
                //rform.Click(".//*[@id='P_L_ClientSubmissionWithAddressAndPOBox_Client2InsName_BirthDate_BirthDate_B-1']");
                //rform.Click(".//*[@id='P_L_ClientSubmissionWithAddressAndPOBox_Client2InsName_BirthDate_BirthDate_DDD_C_T']");
                //ageUtility.Calendar(row, 30, sheetname);
                Console.WriteLine("Policyholder2 details entered");
                Perform.test.Log(LogStatus.Info, "Policyholder2 Details Entered");
            }
            //Console.WriteLine(savescreenshot);
            Perform.ScreenShot(savescreenshot + "Policyholder.png");
            Perform.Click(".//*[@id='ContinueInsImageButtonMiddle']/a");
            System.Threading.Thread.Sleep(2000);
        } 






        
        [Test]
        public void TransactionPage(int row,string sheetname,string savescreenshot)
        {
            PolicyholderPage(row, "New Client Entry", savescreenshot);
            Perform.test = Perform.report.StartTest("Transaction Information Page");
            //Enter transaction info

            Perform.waitTillElementToAppear(".//*[@id='P_L_Application_C0044_VersionEffectiveDateInsDateTime_VersionEffectiveDateInsDateTime']");
            Perform.EnterTextFocus(".//*[@id='P_L_Application_C0044_VersionEffectiveDateInsDateTime_VersionEffectiveDateInsDateTime']", ExcelUtil.GetCellData(row,0, sheetname));
            Console.WriteLine("Date entered");
            Perform.test.Log(LogStatus.Info, "Effective Date Entered");
            Perform.EnterText(".//*[@id='P_L_Application_C0044_LobInsCombo_D_I']", ExcelUtil.GetCellData(row, 1, sheetname));
            Perform.test.Log(LogStatus.Info, "LOB is selected");
            Perform.driver.FindElement(By.XPath(".//*[@id='P_L_Application_C0044_LobInsCombo_D_I']")).SendKeys(Keys.Tab);
            System.Threading.Thread.Sleep(5000);
           



            //Agency Info
            //Perform.SelectDropDown(".//*[@id='P_L_Application_C0044_AgencyInsCombo_D_I']", ExcelUtil.GetCellData(row, 2, sheetname));
            Perform.EnterTextMove(".//*[@id='P_L_Application_C0044_AgencyInsCombo_D_I']", ExcelUtil.GetCellData(row, 2, sheetname));
            Perform.test.Log(LogStatus.Info, "Agency is selected");
            //Perform.driver.FindElement(By.XPath(".//*[@id='P_L_Application_C0044_AgencyInsCombo_D_I']")).Submit();

            System.Threading.Thread.Sleep(5000);
            Perform.ClickPerform(".//*[@id='P_L_Application_C0044_AgencyInsCombo_D_DDD_L_LBI0T0']/em");
            Console.WriteLine("Agency Code Entered");
            System.Threading.Thread.Sleep(2000);
            Perform.ScreenShot(savescreenshot + "TransactionInfo.png");
            Perform.test.Log(LogStatus.Info, "Transaction Information is given");
            Perform.ClickPerform(".//*[@id='SubmitToolStripButtonMiddle']/a");
           // Perform.PageContains("Policyholder Information");
        }
        [Test]
        public void PolicyLevelInfoPage(int row,string sheetname,string savescreenshot)
        {
            TransactionPage(row, "Transaction Information", savescreenshot);
            Perform.test = Perform.report.StartTest("Policy Level Information Page");
            System.Threading.Thread.Sleep(1000);
           // Perform.waitTillElementToAppear(".//*[@id='P_L_V_DetailTreeViewn0Nodes']/table[2]/tbody/tr/td[3]/span/a[2]");
            Perform.Click("//table[2]/tbody/tr/td[3]/span/a[2]");
            Perform.test.Log(LogStatus.Info, "Enter Policy Level Info");
            if (ExcelUtil.GetCellData(row,0,sheetname)!="" && ExcelUtil.GetCellData(row, 0, sheetname) == "YES")
            {
                Perform.Click(".//*[@id='P_L_V_v33w9_t14_c0w0_t0_SelectMarketCreditCheckBox']");
            }
            if (ExcelUtil.GetCellData(row, 1, sheetname) != "" && ExcelUtil.GetCellData(row, 1, sheetname) == "YES")
            {
                Perform.Click(".//*[@id='P_L_V_v33w9_t14_c0w0_t0_AutoHomeInsCheckbox']");
            }
            if (ExcelUtil.GetCellData(row, 2, sheetname) != "" && ExcelUtil.GetCellData(row, 2, sheetname) == "YES")
            {
                Perform.Click(".//*[@id='P_L_V_v33w9_t14_c0w0_t0_EmployeeDiscountInsCheckbox']");
            }
            if (ExcelUtil.GetCellData(row, 3, sheetname) != "" && ExcelUtil.GetCellData(row, 3, sheetname) == "YES")
            {
                Perform.Click(".//*[@id='P_L_V_v33w9_t14_c0w0_t0_FacultativeInsCheckbox']");
            }
            System.Threading.Thread.Sleep(500);
            Perform.test.Log(LogStatus.Info, "Policy level Info is given");
            Perform.ScreenShot(savescreenshot + "PolicyLevelInfo.png");
         
        }
        [Test]
        public void DriverPage(int row, string sheetname, string savescreenshot)
        {
            PolicyLevelInfoPage(row, "Policy_level_Information", savescreenshot);
            Perform.test = Perform.report.StartTest("Driver Page");
            //TransactionPage(row, "Transaction Information", savescreenshot);
            System.Threading.Thread.Sleep(200);
           // Perform.waitTillElementToAppear(".//*[@id='P_L_V_DetailTreeViewn0Nodes']/table[3]/tbody/tr/td[3]/span/a[2]");
            Perform.Click(".//*[@id='P_L_V_DetailTreeViewn0Nodes']/table[3]/tbody/tr/td[3]/span/a[2]");
           // Perform.IsElementPresent(".//*[@id='AddDriverToolStripButtonMiddle']/a");
            Perform.test.Log(LogStatus.Info, "Add New Driver");
            String nofdriver = ExcelUtil.GetCellData(row, 2, "TestCase");
            int col = 0;
            for (int i = 0; i < Int32.Parse(nofdriver); i++)
            {
                if (i > 0)
                {
                    Perform.Click(".//*[@id='AddDriverToolStripButtonMiddle']/a");
                }
                //Enter Driver details
                Perform.SelectTextDropDown(".//*[@id='P_L_V_v33w9_t15_c0w0_PC_t1i" + i + "_c0w0_t0_InsName_Prefix_D_I']", ExcelUtil.GetCellData(row, col, sheetname));
                Perform.EnterText(".//*[@id='P_L_V_v33w9_t15_c0w0_PC_t1i" + i + "_c0w0_t0_InsName_First']", ExcelUtil.GetCellData(row, col + 1, sheetname));
                Perform.EnterText(".//*[@id='P_L_V_v33w9_t15_c0w0_PC_t1i" + i + "_c0w0_t0_InsName_Middle']", ExcelUtil.GetCellData(row, col + 2, sheetname));
                Perform.EnterText(".//*[@id='P_L_V_v33w9_t15_c0w0_PC_t1i" + i + "_c0w0_t0_InsName_Last']", ExcelUtil.GetCellData(row, col + 3, sheetname));
                Console.WriteLine("Last name entered");
                // Perform.SelectTextDropDown(".//*[@id='P_L_V_v33w9_t15_c0w0_PC_t1i"+i+"_c0w0_t0_InsName_Suffix_D_I']", ExcelUtil.GetCellData(row, col+4, sheetname));
                //if (i > 0)
               // {
                    Perform.EnterText(".//*[@id='P_L_V_v33w9_t15_c0w0_PC_t1i" + i + "_c0w0_t0_InsName_PersonalTaxNumber_PersonalTaxNumber']", ExcelUtil.GetCellData(row, col + 5, sheetname));
                    Console.WriteLine("SSN entered");
              //  }
                Perform.Click(".//*[@id='P_L_V_v33w9_t15_c0w0_PC_t1i" + i + "_c0w0_t0_InsName_Sex_D_I']");

                string gender = ExcelUtil.GetCellData(row, col + 6, sheetname);

                if (gender == "Male")
                {
                    Console.WriteLine(gender);
                    Perform.Click(".//*[@id='P_L_V_v33w9_t15_c0w0_PC_t1i" + i + "_c0w0_t0_InsName_Sex_D_DDD_L_LBI1T1']");
                }
                if (gender == "Female")
                {
                    Console.WriteLine(gender);
                    Perform.Click(".//*[@id='P_L_V_v33w9_t15_c0w0_PC_t1i" + i + "_c0w0_t0_InsName_Sex_D_DDD_L_LBI2T1']");
                }
                Console.WriteLine("Gender is given");

                string maritalstatus = ExcelUtil.GetCellData(row, col + 7, sheetname);

                Perform.Click(".//*[@id='P_L_V_v33w9_t15_c0w0_PC_t1i" + i + "_c0w0_t0_InsName_MaritalStatus_D_I']");
                System.Threading.Thread.Sleep(200);
                if (maritalstatus == "Divorced")
                {
                    Perform.Click(".//*[@id='P_L_V_v33w9_t15_c0w0_PC_t1i" + i + "_c0w0_t0_InsName_MaritalStatus_D_DDD_L_LBI1T1']");
                }
                if (maritalstatus == "Married")
                {
                    Perform.Click(".//*[@id='P_L_V_v33w9_t15_c0w0_PC_t1i" + i + "_c0w0_t0_InsName_MaritalStatus_D_DDD_L_LBI2T1']");
                }
                if (maritalstatus == "Single")
                {
                    Perform.Click(".//*[@id='P_L_V_v33w9_t15_c0w0_PC_t1i" + i + "_c0w0_t0_InsName_MaritalStatus_D_DDD_L_LBI3T1']");
                }
                if (maritalstatus == "Widowed")
                {
                    Perform.Click(".//*[@id='P_L_V_v33w9_t15_c0w0_PC_t1i" + i + "_c0w0_t0_InsName_MaritalStatus_D_DDD_L_LBI4T1']");
                }
                System.Threading.Thread.Sleep(500);
                Perform.EnterText(".//*[@id='P_L_V_v33w9_t15_c0w0_PC_t1i" + i + "_c0w0_t0_InsName_LicenseNumber_LicenseNumber']", ExcelUtil.GetCellData(row, col + 8, sheetname));
                System.Threading.Thread.Sleep(200);
                
                //Perform.SelectTextDropDown(".//*[@id='P_L_V_v33w9_t15_c0w0_PC_t1i" + i + "_c0w0_t0_InsName_LicenseCountry_D_I']", ExcelUtil.GetCellData(row, col + 9, sheetname));

                Perform.SelectTextDropDown(".//*[@id='P_L_V_v33w9_t15_c0w0_PC_t1i" + i + "_c0w0_t0_InsName_LicenseState_D_I']", ExcelUtil.GetCellData(row, col + 10, sheetname));
                System.Threading.Thread.Sleep(200);
                /* if(i>0)
                 { 
                 Perform.Click("//img[@id='P_L_V_v33w9_t15_c0w0_PC_t1i"+i+"_c0w0_t0_InsName_BirthDate_BirthDate_B-1Img']");
                     System.Threading.Thread.Sleep(500);
                     Perform.Click(".//td[@id='P_L_V_v33w9_t15_c0w0_PC_t1i"+i+"_c0w0_t0_InsName_BirthDate_BirthDate_DDD_C_TC']/span");


                     PageUtility.Calendardriver(row,col+11,i, sheetname);
                 }*/
                Perform.EnterTextFocus(".//*[@id='P_L_V_v33w9_t15_c0w0_PC_t1i"+i+"_c0w0_t0_InsName_BirthDate_BirthDate_I']", ExcelUtil.GetCellData(row, col + 11, sheetname));

                if (ExcelUtil.GetCellData(row, col + 15, sheetname) != "")
                {
                    Perform.EnterText(".//*[@id='P_L_V_v33w9_t15_c0w0_PC_t1i" + i + "_c0w0_t0_  MotorcycleTrainingMembershipInsDateTime_MotorcycleTrainingMembershipInsDateTime']", ExcelUtil.GetCellData(row, col + 13, sheetname));
                    if (ExcelUtil.GetCellData(row, col + 16, sheetname)!="" && ExcelUtil.GetCellData(row, col + 16, sheetname) == "YES")
                    {
                        Perform.Click(".//*[@id='P_L_V_v33w9_t15_c0w0_PC_t1i" + i + "_c0w0_t0_MotorcycleTrainingDiscountCheckBox']");
                    }
                }
                else
                {
                    Console.WriteLine("No Motorcycle Details");
                }
                /* Perform.driver.FindElement(By.XPath("//*[@id='P_L_V_v33w9_t15_c0w0_PC_t1i" + i + "_c0w0_t0_RelationToPolicyHolderInsCombo_D']")).Click();
                 IList<IWebElement> rel = Perform.driver.FindElements(By.XPath("//td[contains(@id,'P_L_V_v33w9_t15_c0w0_PC_t1i" + i + "_c0w0_t0_RelationToPolicyHolderInsCombo_D_DDD_L')]"));
                 System.Threading.Thread.Sleep(500);*/
                Console.WriteLine(ExcelUtil.GetCellData(row, col + 17, sheetname));
               // Perform.Click(".//*[@id='P_L_V_v33w9_t15_c0w0_PC_t1i"+i+"_c0w0_t0_RelationToPolicyHolderInsCombo_D_B-1Img']");
                System.Threading.Thread.Sleep(200);
                // PageUtility.policyrelation(row, i, col + 17, sheetname);
                if (i > 0)
                {
                    Perform.EnterTextFocus(".//*[@id='P_L_V_v33w9_t15_c0w0_PC_t1i" + i + "_c0w0_t0_RelationToPolicyHolderInsCombo_D_I']", ExcelUtil.GetCellData(row, col + 17, sheetname));
                }



                Perform.SelectListElement("//*[@id='P_L_V_v33w9_t15_c0w0_PC_t1i" + i + "_c0w0_t0_RatedExcludedInsCombo_D_I']", ExcelUtil.GetCellData(row, col + 18, sheetname));
                
                Perform.EnterText(".//*[@id='P_L_V_v33w9_t15_c0w0_PC_t1i" + i + "_c0w0_t0_ReasonExcludedInsTextBox']", ExcelUtil.GetCellData(row, col + 19, sheetname));
                if (ExcelUtil.GetCellData(row, col + 20, sheetname) != "" && ExcelUtil.GetCellData(row, col + 20, sheetname) == "YES")
                {
                    Perform.Click(".//*[@id='P_L_V_v33w9_t15_c0w0_PC_t1i" + i + "_c0w0_t0_DistantStudentInsCheckBox']");
                }
                if (ExcelUtil.GetCellData(row, col + 21, sheetname) != "" && ExcelUtil.GetCellData(row, col + 21, sheetname) == "YES")
                {
                    Perform.Click(".//*[@id='P_L_V_v33w9_t15_c0w0_PC_t1i" + i + "_c0w0_t0_GoodStudentInsCheckBox']");
                }
                if (ExcelUtil.GetCellData(row, col + 22, sheetname) != "")
                {
                    Perform.EnterText(".//*[@id='P_L_V_v33w9_t15_c0w0_PC_t1i" + i + "_c0w0_t0_DefensiveDriverInsDateTime_DefensiveDriverInsDateTime']", ExcelUtil.GetCellData(row, col + 22, sheetname));
                }
                if (ExcelUtil.GetCellData(row, col + 23, sheetname) != "")
                {
                    Perform.EnterText(".//*[@id='P_L_V_v33w9_t15_c0w0_PC_t1i" + i + "_c0w0_t0_AccidentPreventionCourseInsDateTime_AccidentPreventionCourseInsDateTime']", ExcelUtil.GetCellData(row, col + 23, sheetname));
                }
                Perform.EnterText(".//*[@id='P_L_V_v33w9_t15_c0w0_PC_t1i" + i + "_c0w0_t0_DistanceToSchoolInsNumeric_DistanceToSchoolInsNumeric']", ExcelUtil.GetCellData(row, col + 24, sheetname));
                col = col + 25;
                Perform.Click(".//*[@id='SaveToolStripButtonMiddle']/a");
                Perform.test.Log(LogStatus.Info, "Driver General Information Entered");

                // if (i = Int32.Parse(nofdriver) - 1)

                //Extended Non-Owned Information
                string sheetname1 = "Extended Non-Owned";
                if (ExcelUtil.GetCellData(row, 0, sheetname1) != "")
                {
                    Perform.Click(".//*[@id='P_L_V_v33w9_t15_c0w0_PC_T3T']/span");
                    if (ExcelUtil.GetCellData(row, 0, sheetname1) == "YES")
                    {
                        Perform.Click(".//*[@id='P_L_V_v33w9_t15_c0w0_PC_t3i0_ExtendedNonOwnedInsCheckbox']");
                    }
                    if (ExcelUtil.GetCellData(row, 1, sheetname1) == "YES")
                    {
                        Perform.Click(".//*[@id='P_L_V_v33w9_t15_c0w0_PC_t3i0_PrimaryLiabilityInsuranceProvidedInsCheckbox']");
                    }
                    if (ExcelUtil.GetCellData(row, 2, sheetname1) == "YES")
                    {
                        Perform.Click(".//*[@id='P_L_V_v33w9_t15_c0w0_PC_t3i0_AutoUsedInGovernmentBusinessInsCheckbox']");
                    }
                    if (ExcelUtil.GetCellData(row, 3, sheetname1) == "YES")
                    {
                        Perform.Click(".//*[@id='P_L_V_v33w9_t15_c0w0_PC_t3i0_RegularUseInsCheckbox']");
                    }
                    if (ExcelUtil.GetCellData(row, 4, sheetname1) == "YES")
                    {
                        Perform.Click(".//*[@id='P_L_V_v33w9_t15_c0w0_PC_t3i0_EmployedByGarageInsCheckbox']");
                    }
                    if (ExcelUtil.GetCellData(row, 5, sheetname1) == "Named Individuals and Resident Relatives")
                    {
                        Perform.Click(".//*[@id='P_L_V_v33w9_t15_c0w0_PC_t3i0_NamedInsuredResidentSpouseRadioButton']");
                    }
                    if (ExcelUtil.GetCellData(row, 5, sheetname1) == "Named Individuals Only")
                    {
                        Perform.Click(".//*[@id='P_L_V_v33w9_t15_c0w0_PC_t3i0_NamedInsuredRadioButton']");
                    }
                }
                Perform.test.Log(LogStatus.Info, "Driver Extended Non-Owned Info Entered");

            }

            //Accident Violations





            System.Threading.Thread.Sleep(500);
            Perform.test.Log(LogStatus.Info, "Driver Details Entered");

            Perform.ScreenShot(savescreenshot + "Driver.png");
        }
        [Test]
        public void VehiclePage(int row,string sheetname,string savescreenshot)
        {
            DriverPage(row, "Driver Page", savescreenshot);
            Perform.test = Perform.report.StartTest("Vehicle Page");
            System.Threading.Thread.Sleep(1000);
         
            Perform.Click(".//*[@id='P_L_V_DetailTreeViewn0Nodes']/table[4]/tbody/tr/td[3]/span/a[2]");
            Console.WriteLine("Vehicles Link is clicked");
          
            Perform.waitTillElementToAppear(".//*[@id='AddVehicleImageButtonMiddle']/a");
            Perform.IsElementPresent(".//*[@id='AddVehicleImageButtonMiddle']/a");
            string nofvehicle = ExcelUtil.GetCellData(row, 3,"TestCase");
            string nofdriver = ExcelUtil.GetCellData(row, 2, "TestCase");
            int col = 0;
            Perform.test.Log(LogStatus.Info, "Enter Vehicle Details");
            for (int i = 0; i < Int32.Parse(nofvehicle); i++)
            {
                int ocd = col + 7;
                Perform.Click(".//*[@id='AddVehicleImageButtonMiddle']/a");
               Perform.EnterText(".//*[@id='P_L_V_v33w9_t16_c0w0_PC_t1i" + i + "_YearInsNumeric_YearInsNumeric']", ExcelUtil.GetCellData(row, col, sheetname));
                System.Threading.Thread.Sleep(200);
               Perform.EnterTextMove(".//*[@id='P_L_V_v33w9_t16_c0w0_PC_t1i" + i + "_MakeInsTextBox']", ExcelUtil.GetCellData(row, col + 1, sheetname));
               Perform.EnterText(".//*[@id='P_L_V_v33w9_t16_c0w0_PC_t1i" + i + "_ModelInsTextBox']", ExcelUtil.GetCellData(row, col + 2, sheetname));
               Perform.SelectTextDropDown(".//*[@id='P_L_V_v33w9_t16_c0w0_PC_t1i" + i + "_BodyTypeInsCombo_D_I']", ExcelUtil.GetCellData(row, col + 3, sheetname));
                Perform.EnterText(".//*[@id='P_L_V_v33w9_t16_c0w0_PC_t1i" + i + "_VINInsTextBox_I']", ExcelUtil.GetCellData(row, col + 4, sheetname));
                Perform.Click(".//*[@id='CustomAction0ToolStripButtonMiddle']/a");
                Perform.Click(".//*[@id='VinLookupButtonMiddle']/a");
                Perform.Click(".//*[@id='P_L_V_v33w9_t16_c0w0_PC_t0_ModelIsoDataGridView_SelectModelLinkButton_0']");
                Perform.EnterText(".//*[@id='P_L_V_v33w9_t16_c0w0_PC_t1i" + i + "_CostNewInsNumeric_CostNewInsNumeric']", ExcelUtil.GetCellData(row, col + 5, sheetname));
                System.Threading.Thread.Sleep(1000);
                Perform.test.Log(LogStatus.Info, "Basic Vehicle Info Given");
                Perform.SelectTextDropDown(".//*[@id='P_L_V_v33w9_t16_c0w0_PC_t1i" + i + "_PrincipalDriverInsCombo_D_I']", ExcelUtil.GetCellData(row, col + 6, sheetname));
                for (int j = 1; j < Int32.Parse(nofdriver); j++)
                {
                    if (ExcelUtil.GetCellData(row, ocd, sheetname) != "")
                    {
                        System.Threading.Thread.Sleep(200);
                        Perform.EnterTextTab(".//*[@id='P_L_V_v33w9_t16_c0w0_PC_t1i" + i + "_OccasionalDriver" + j + "InsCombo_D_I']", ExcelUtil.GetCellData(row, ocd, sheetname));
                        
                    }
                    ocd = ocd + 1;

                }
                Perform.test.Log(LogStatus.Info, "Driver Assignment Given");
                if (ExcelUtil.GetCellData(row, col + 10, sheetname) == "YES")
                {
                    Perform.Click(".//*[@id='P_L_V_v33w9_t16_c0w0_PC_t1i" + i + "_MultiCarInsCheckbox']");
                }
                else
                {
                    Console.WriteLine("No Multi car option");
                }
                Perform.test.Log(LogStatus.Info, "Discount and Surcharge Info Given");
                if (ExcelUtil.GetCellData(row, col + 11, sheetname) != "")
                {
                    Perform.SelectDropDown(".//*[@id='P_L_V_v33w9_t16_c0w0_PC_t1i" + i + "_VehicleTypeInsCombo_D_I']", ExcelUtil.GetCellData(row, col + 11, sheetname));
                    Perform.EnterText(".//*[@id='P_L_V_v33w9_t16_c0w0_PC_t1i" + i + "_CubicCentimetersInsNumeric_CubicCentimetersInsNumeric']", ExcelUtil.GetCellData(row, col + 12, sheetname));
                    Perform.EnterText(".//*[@id='P_L_V_v33w9_t16_c0w0_PC_t1i" + i + "_StatedAmountInsNumeric_StatedAmountInsNumeric']", ExcelUtil.GetCellData(row, col + 13, sheetname));
                }
                else
                {
                    Console.WriteLine("No Motorcycle Info");
                }
                Perform.test.Log(LogStatus.Info, "Motorcycle Info Given");
                Perform.Click(".//*[@id='SaveToolStripButtonMiddle']/a");
                col = col + 14;
            }
            Perform.test.Log(LogStatus.Info, "Vehicle General Information Given");
            //Additional Interest
            /* string sheetname1 = "Additional_Interest";
             if(ExcelUtil.GetCellData(row,0,sheetname1)!="")
             {
                 Perform.Click(".//*[@id='P_L_V_v33w9_t16_c0w0_PC_AT2T']/span");
                 Perform.Click(".//*[@id='AddAdditionalInterestImageButtonMiddle']/a");

             }*/
            System.Threading.Thread.Sleep(500);
            Perform.test.Log(LogStatus.Info, "Vehicle Details Given");
            Perform.ScreenShot(savescreenshot + "Vehicle.png");
              
            
        }
        [Test]
        public void AdditionalPolicyInfoPage(int row,string sheetname,string savescreenshot)
        {
            VehiclePage(row, "Vehicles", savescreenshot);
            Perform.test = Perform.report.StartTest("Additional Policy Info Page");
            System.Threading.Thread.Sleep(1000);
            Perform.Click("//a[contains(text(),'Additonal Policy Info')]");
            Perform.test.Log(LogStatus.Info, "Select 3rd Party Report Type");
            Perform.SelectDropDown(".//*[@id='P_L_V_v33w9_t17_c0w0_NB_ITC0i0_t0_0_AvailableProductsListBox_0']","MVR");
            Perform.Click(".//*[@id='SelectProductButtonMiddle']/a");
            System.Threading.Thread.Sleep(500);
            Perform.SelectDropDown(".//*[@id='P_L_V_v33w9_t17_c0w0_NB_ITC0i0_t0_0_AvailableProductsListBox_0']","CLUE");
            Perform.Click(".//*[@id='SelectProductButtonMiddle']/a");
            System.Threading.Thread.Sleep(500);
            Perform.SelectDropDown(".//*[@id='P_L_V_v33w9_t17_c0w0_NB_ITC0i0_t0_0_AvailableProductsListBox_0']","Credit Report");
            Perform.Click(".//*[@id='SelectProductButtonMiddle']/a");
            System.Threading.Thread.Sleep(1000);
            Perform.Click(".//span[contains(@id,'checkAllSubjectsCheckbox_0_S_D')]");
            Perform.Click(".//*[@id='OrderToolstripButtonMiddle']/a");
            Perform.test.Log(LogStatus.Info, "Reports Ordered");
            Perform.Click(".//*[@id='P_L_V_v33w9_t17_c0w0_NB_ITC0i0_t0_0_ChoicepointControlValidationList_0_MyASPxPopupControl_0_ContinueInsValidationButton_0_CD']/span");
            Perform.test.Log(LogStatus.Warning, "Continue button is clicked");
            Console.WriteLine("Continue button is clicked");
            Perform.waitTillElementToAppear(".//*[@id='P_L_V_ValidationPopUp_MyASPxPopupControl_OtherLabel']");
          //  Perform.driver.SwitchTo().Alert().Accept();
           Perform.Click(".//*[@id='P_L_V_ValidationPopUp_MyASPxPopupControl_OKInsValidationButton_CD']/span");
            Perform.test.Log(LogStatus.Info, "OK button is clicked in Message Box");
            System.Threading.Thread.Sleep(500);
            Perform.test.Log(LogStatus.Info, "Additional POlicy Info is Given");
            Perform.ScreenShot(savescreenshot + "AdditionalPolicyInfo.png");
           // Perform.Click(".//*[@id='P_L_V_RateToolStripButton']");
        }
        [Test]
        public void CoveragePage(int row, string sheetname, string savescreenshot)
        {
            AdditionalPolicyInfoPage(row, "Additional_Policy_Info", savescreenshot);
              Perform.test = Perform.report.StartTest("Coverage Page");
            System.Threading.Thread.Sleep(500);
            Perform.Click(".//*[@id='P_L_V_DetailTreeViewn0Nodes']/table[6]/tbody/tr/td[3]/span/a[2]");
            Perform.test.Log(LogStatus.Info, "Enter Vehicle Coverage Details");
            System.Threading.Thread.Sleep(500);
            if (ExcelUtil.GetCellData(row, 0, sheetname) == "YES")
            {
                Perform.Click(".//*[@id='P_L_V_v33w9_t18_c0w0_t0_PolicyLevelDynamicCoverageTableLayoutPanel_13105_80443_13105_80443_MainLimitLimit']");
            }
            if (ExcelUtil.GetCellData(row, 1, sheetname) == "YES")
            {
                Perform.Click(".//*[@id='P_L_V_v33w9_t18_c0w0_t0_PolicyLevelDynamicCoverageTableLayoutPanel_13104_80094_13104_80094_MainLimitLimit']");
            }
            System.Threading.Thread.Sleep(200);
            Perform.EnterTextTab(".//*[@id='P_L_V_v33w9_t18_c0w0_t0_VehicleLevelSameLimitsDynamicCoverageTableLayoutPanel_13086_2_13086_2_MainLimitLimit_D_I']", ExcelUtil.GetCellData(row, 2, sheetname));
            System.Threading.Thread.Sleep(200);
            Perform.EnterTextTab(".//*[@id='P_L_V_v33w9_t18_c0w0_t0_VehicleLevelSameLimitsDynamicCoverageTableLayoutPanel_13085_1_13085_1_MainLimitLimit_D_I']", ExcelUtil.GetCellData(row, 3, sheetname));
            System.Threading.Thread.Sleep(200);
            //Perform.driver.FindElement(By.XPath(".//*[@id='P_L_V_v33w9_t18_c0w0_t0_VehicleLevelSameLimitsDynamicCoverageTableLayoutPanel_13085_1_13085_1_MainLimitLimit_D_I']")).SendKeys(Keys.Tab);
            Perform.EnterTextTab(".//*[@id='P_L_V_v33w9_t18_c0w0_t0_VehicleLevelSameLimitsDynamicCoverageTableLayoutPanel_13088_4_13088_4_MainLimitLimit_D_I']", ExcelUtil.GetCellData(row, 4, sheetname));
            System.Threading.Thread.Sleep(200);
            Perform.EnterTextTab(".//*[@id='P_L_V_v33w9_t18_c0w0_t0_VehicleLevelSameLimitsDynamicCoverageTableLayoutPanel_13090_6_13090_6_MainLimitLimit_D_I']", ExcelUtil.GetCellData(row, 5, sheetname));
            System.Threading.Thread.Sleep(200);
            Perform.EnterTextTab(".//*[@id='P_L_V_v33w9_t18_c0w0_t0_VehicleLevelSameLimitsDynamicCoverageTableLayoutPanel_13099_10007_13099_10007_MainLimitLimit_D_I']", ExcelUtil.GetCellData(row, 6, sheetname));
            System.Threading.Thread.Sleep(200);
            Perform.EnterTextTab(".//*[@id='P_L_V_v33w9_t18_c0w0_t0_VehicleLevelSameLimitsDynamicCoverageTableLayoutPanel_13091_8_13091_8_MainLimitLimit_D_I']", ExcelUtil.GetCellData(row, 7, sheetname));
            System.Threading.Thread.Sleep(200);
            Perform.EnterTextTab(".//*[@id='P_L_V_v33w9_t18_c0w0_t0_VehicleLevelSameLimitsDynamicCoverageTableLayoutPanel_13092_9_13092_9_MainLimitLimit_D_I']", ExcelUtil.GetCellData(row, 8, sheetname));
            System.Threading.Thread.Sleep(200);
            Perform.EnterTextTab(".//*[@id='P_L_V_v33w9_t18_c0w0_t0_VehicleLevelSameLimitsDynamicCoverageTableLayoutPanel_13097_293_13097_293_MainLimitLimit_D_I']", ExcelUtil.GetCellData(row, 9, sheetname));
            Perform.test.Log(LogStatus.Info, "General Vehicle Level Coverage Entered");
            //Select Vehicle
            string nofvehicle = ExcelUtil.GetCellData(row, 3, "TestCase");
            int vcol = 11;
            for (int i = 0; i < Int32.Parse(nofvehicle); i++)
            {
                Perform.Click(".//*[@id='P_L_V_v33w9_t18_c0w0_t0_VehicleInsDataGridView']/tbody/tr[" + (i + 2) + "]/td[1]/a");
                if (ExcelUtil.GetCellData(row, vcol, sheetname) == "YES")
                {
                    Perform.Click(".//*[@id='P_L_V_v33w9_t18_c0w0_t0_ComprehensiveOnlyInsCheckbox']");
                }
                System.Threading.Thread.Sleep(500);
                // Perform.Wait();
                // Perform.waitTillElementToAppear(".//*[@id='P_L_V_v33w9_t18_c0w0_t0_VehicleLevelDynamicCoverageTableLayoutPanel_13087_3_13087_3_MainLimitLimit_D_I']");
                Perform.EnterTextFocus(".//*[@id='P_L_V_v33w9_t18_c0w0_t0_VehicleLevelDynamicCoverageTableLayoutPanel_13087_3_13087_3_MainLimitLimit_D_I']", ExcelUtil.GetCellData(row, vcol + 1, sheetname));
             
                // Perform.driver.FindElement(By.XPath(".//*[@id='P_L_V_v33w9_t18_c0w0_t0_VehicleLevelDynamicCoverageTableLayoutPanel_13087_3_13087_3_MainLimitLimit_D_I']")).Clear();
               // Perform.EnterTextTab(".//*[@id='P_L_V_v33w9_t18_c0w0_t0_VehicleLevelDynamicCoverageTableLayoutPanel_13087_3_13087_3_MainLimitLimit_D_I']", ExcelUtil.GetCellData(row, vcol + 1, sheetname));
                Perform.waitTillElementToAppear(".//*[@id='P_L_V_v33w9_t18_c0w0_t0_VehicleLevelDynamicCoverageTableLayoutPanel_13089_5_13089_5_MainLimitLimit_D_I']");
                
                Perform.EnterTextTab(".//*[@id='P_L_V_v33w9_t18_c0w0_t0_VehicleLevelDynamicCoverageTableLayoutPanel_13089_5_13089_5_MainLimitLimit_D_I']", ExcelUtil.GetCellData(row, vcol + 2, sheetname));
                if (ExcelUtil.GetCellData(row, vcol + 3, sheetname) == "YES")
                {
                    Perform.Click(".//*[@id='P_L_V_v33w9_t18_c0w0_t0_VehicleLevelDynamicCoverageTableLayoutPanel_13103_80056_13103_80056_MainLimitLimit']");
                }
                System.Threading.Thread.Sleep(500);
                Perform.EnterTextTab(".//*[@id='P_L_V_v33w9_t18_c0w0_t0_VehicleLevelDynamicCoverageTableLayoutPanel_13101_60008_13101_60008_MainLimitLimit_D_I']", ExcelUtil.GetCellData(row, vcol + 4, sheetname));
                Perform.waitTillElementToAppear(".//*[@id='P_L_V_v33w9_t18_c0w0_t0_VehicleLevelDynamicCoverageTableLayoutPanel_13096_66_13096_66_MainLimitLimit_D_I']");
                Perform.EnterTextTab(".//*[@id='P_L_V_v33w9_t18_c0w0_t0_VehicleLevelDynamicCoverageTableLayoutPanel_13096_66_13096_66_MainLimitLimit_D_I']", ExcelUtil.GetCellData(row, vcol + 5, sheetname));
                if (ExcelUtil.GetCellData(row, vcol + 6, sheetname) == "YES")
                {
                    Perform.Click(".//*[@id='P_L_V_v33w9_t18_c0w0_t0_VehicleLevelDynamicCoverageTableLayoutPanel_13100_10044_13100_10044_MainLimitLimit']");
                }
                Perform.EnterTextTab(".//*[@id='P_L_V_v33w9_t18_c0w0_t0_VehicleLevelDynamicCoverageTableLayoutPanel_13094_16_13094_16_MainLimitLimit_D_I']", ExcelUtil.GetCellData(row, vcol + 7, sheetname));
                Perform.EnterText(".//*[@id='P_L_V_v33w9_t18_c0w0_t0_VehicleLevelDynamicCoverageTableLayoutPanel_13093_15_13093_15_MainLimitLimit_13093_15_MainLimitLimit_I']", ExcelUtil.GetCellData(row, vcol + 8, sheetname));
                Perform.EnterText(".//*[@id='P_L_V_v33w9_t18_c0w0_t0_VehicleLevelDynamicCoverageTableLayoutPanel_13095_57_13095_57_MainLimitLimit_13095_57_MainLimitLimit_I']", ExcelUtil.GetCellData(row, vcol + 9, sheetname));
                Perform.EnterTextTab(".//*[@id='P_L_V_v33w9_t18_c0w0_t0_VehicleLevelDynamicCoverageTableLayoutPanel_13102_80031_13102_80031_MainLimitLimit_D_I']", ExcelUtil.GetCellData(row, vcol + 10, sheetname));
                vcol = vcol + 12;
            }
            Perform.test.Log(LogStatus.Info, "Specific Vehicle Level Coverage Entered");
            Perform.ScreenShot(savescreenshot + "Coverage.png");
        }
        [Test]
        public void BillingPage(int row, string sheetname, string savescreenshot)
        {
            CoveragePage(row, "Coverages", savescreenshot);
            Perform.test = Perform.report.StartTest("Billing Page");
            System.Threading.Thread.Sleep(1000);
            Perform.Click(".//*[@id='P_L_V_DetailTreeViewn0Nodes']/table[7]/tbody/tr/td[3]/span/a[2]");
            Perform.SelectTextDropDown(".//*[@id='P_L_V_v33w9_t19_MethodComboBox_D_I']", ExcelUtil.GetCellData(row, 0, sheetname));
            System.Threading.Thread.Sleep(1000);
            Perform.SelectTextDropDown(".//*[@id='P_L_V_v33w9_t19_PayPlanInsCombo_D_I']", ExcelUtil.GetCellData(row, 1, sheetname));
            System.Threading.Thread.Sleep(1000);
            Perform.SelectTextDropDown(".//*[@id='P_L_V_v33w9_t19_BillToControl_BillToInsComboBox_D_I']", ExcelUtil.GetCellData(row, 2, sheetname));
            Console.WriteLine("Billing Info is given");
            Perform.test.Log(LogStatus.Info, "Billing Details Entered");
        }
        [Test]
        public void UnderwritingPage(int row, string sheetname, string savescreenshot)
        {
            BillingPage(row, "Billing", savescreenshot);
            Perform.test = Perform.report.StartTest("Underwriting Page");
            System.Threading.Thread.Sleep(1000);
            Perform.Click(".//*[@id='P_L_V_DetailTreeViewn0Nodes']/table[8]/tbody/tr/td[3]/span/a[2]");
            Perform.test.Log(LogStatus.Info, "Answer Underwriting Questions");
            int val = 9283;
            int col = 0;
            for (int i=0;i<16;i++)
            {
              
                //if(ExcelUtil.GetCellData(row,col,sheetname)=="YES")
                //{
                  //  Perform.Click(".//*[@id='P_L_V_v33w9_t20_radiobutton_"+val+"_110_1_Yes']");
                    //Perform.EnterText(".//*[@id='P_L_V_v33w9_t20_textbox_" + val + "_110_1_AdditionalInformation']", ExcelUtil.GetCellData(row, col + 1, sheetname));
                //}
                if(ExcelUtil.GetCellData(row, col, sheetname) == "NO")
                {
                    Perform.Click(".//*[@id='P_L_V_v33w9_t20_radiobutton_" + val + "_110_1_No']");
                }
                col = col + 1;
                val = val + 1;
            }
            Perform.test.Log(LogStatus.Info, "Underwriting Questions Answered");
            Perform.ScreenShot(savescreenshot + "Underwriting.png");
        }
        [Test]
        public void ViewQuotePage()
        {
            int sheetrownum = ExcelUtil.getRowCount("TestCase");
            try
            {
                for (int i = 2; i < sheetrownum; i++)
                {
                    System.IO.Directory.CreateDirectory(path + ExcelUtil.GetCellData(i, 0, "TestCase"));
                    string savescreenshot = path + ExcelUtil.GetCellData(i, 0, "TestCase") + "\\";
                    Console.WriteLine(savescreenshot);
                    UnderwritingPage(i, "Underwriting", savescreenshot);
                    Perform.test = Perform.report.StartTest("View Quote Page");
                    //View Quote
                    Perform.Click(".//*[@id='P_L_V_RateToolStripButton']");

                    if (Perform.driver.FindElement(By.XPath(".//*[@id='P_L_V_ValidationPopUp_MyASPxPopupControl_ContinueInsValidationButton_CD']/span")).Displayed)
                    {
                        Perform.Click(".//*[@id='P_L_V_ValidationPopUp_MyASPxPopupControl_ContinueInsValidationButton_CD']/span");
                        Perform.test.Log(LogStatus.Warning, "Continue button Clicked");
                    }
                    System.Threading.Thread.Sleep(1000);
                    Perform.Click(".//*[@id='P_L_V_DetailTreeViewn0Nodes']/table[9]/tbody/tr/td[3]/span/a[2]");
                    Perform.Click("//div[@id='P_L_V_ValidationPopUp_MyASPxPopupControl_OKInsValidationButton_CD']/span");
                    Perform.test.Log(LogStatus.Warning, "OK button Clicked");
                    System.Threading.Thread.Sleep(500);
                    Perform.ScreenShot(savescreenshot + "Policy" + (i - 1) + "_" + DateTime.Now.ToString("yyyyMMddhhmmss") + ".png");
                    Console.WriteLine("Policy" + (i - 1) + "is isssued");
                    Perform.test.Log(LogStatus.Pass, "Policy is Issued");
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("Exception" + e);
                Assert.True(false);
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
