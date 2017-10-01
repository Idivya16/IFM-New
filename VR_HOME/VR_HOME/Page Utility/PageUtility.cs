using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace VR_HOME
{
    class PageUtility
    {
        public static void Coveragedetails(int row,string sheetname)
        {

            //Deductibles
            Perform.waitTillElementToAppear(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ddlDeductible']");
            Perform.SelectDropDown(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ddlDeductible']", ExcelUtil.GetCellData(row,0, sheetname));
            if (ExcelUtil.GetCellData(row,1,sheetname)!="")
            {
                Perform.SelectDropDown(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ddlWindHailDeductible']", ExcelUtil.GetCellData(row, 1, sheetname));
            }
            else { Console.WriteLine("Hail deductible is not displayed"); }
            //Policy Basic Coverages
            if (ExcelUtil.GetCellData(row, 2, sheetname) != "")
            {
                if (ExcelUtil.GetCellData(row, 0, "FORM") == "HO-6 - HOMEOWNERS UNIT OWNERS FORM" || ExcelUtil.GetCellData(row, 0, "FORM") == "HO-6 - HOMEOWNERS UNIT OWNERS FORM")
                {
                    Perform.EnterText(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_txtDwellingChangeInLimit']", ExcelUtil.GetCellData(row, 2, sheetname));
                }
                else
                {
                    Perform.EnterText(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_txtDwLimit']", ExcelUtil.GetCellData(row, 2, sheetname));
                }
            }
            if (ExcelUtil.GetCellData(row, 3, sheetname) != "")
            {
                Perform.EnterText(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_txtRPSChgInLimit']", ExcelUtil.GetCellData(row, 3, sheetname));
            }
            if (ExcelUtil.GetCellData(row, 0, "FORM") == "HO-4 - HOMEOWNERS CONTENTS BROAD FORM" || ExcelUtil.GetCellData(row, 0, "FORM") == "ML-4 - MOBILE HOME TENANT OCCUPIED" || ExcelUtil.GetCellData(row, 0, "FORM") == "HO-6 - HOMEOWNERS UNIT OWNERS FORM")
            {
                Perform.EnterText(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_txtPPLimit']", ExcelUtil.GetCellData(row, 4, sheetname));
            }
            else
            {
                Perform.EnterText(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_txtPPChgInLimit']", ExcelUtil.GetCellData(row, 4, sheetname));
            }
           // if(ExcelUtil.GetCellData(row,5,sheetname)!="")
            //{
              //  Perform.EnterText("", ExcelUtil.GetCellData(row, 5, sheetname));
            //}
            Perform.SelectDropDown(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ddlPersonalLiability']", ExcelUtil.GetCellData(row,6, sheetname));
            Perform.SelectDropDown(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ddlMedicalPayments']", ExcelUtil.GetCellData(row,7, sheetname));
            Perform.Click(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_btnSaveBase']");
            Console.WriteLine("Basic Policy Coverage is entered");
            //Optional Coverages
           string sheetname0 = "Optional_Coverages";
            Perform.driver.FindElement(By.XPath("//label[contains(@id,'lblHomOptionCoverageHeader')]")).Click();

            System.Threading.Thread.Sleep(500);
            
           /* if(ExcelUtil.GetCellData(row,0,sheetname0)=="YES" && ExcelUtil.GetCellData(row, 0, sheetname0)!="")
            {
                Perform.Click("");
            }*/
            //Personal Property Replacement Cost(HO 290/92/195)
           if (ExcelUtil.GetCellData(row,1,sheetname0)!="YES" && ExcelUtil.GetCellData(row, 1, sheetname0)!="")
               {
                   Perform.Click(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlHomSectionCoverages_Repeater1_ctlSectionCoverageItem_7_chkCov_7']");
               }
            // Personal Property Replacement Cost (ML - 55)
            if (ExcelUtil.GetCellData(row,2,sheetname0)=="YES" && ExcelUtil.GetCellData(row, 2, sheetname0) != "" && ExcelUtil.GetCellData(row,0,"FORM")== "ML-2 - MOBILE HOME OWNER OCCUPIED")
            {
                Perform.Click(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlHomSectionCoverages_Repeater1_ctlSectionCoverageItem_7_chkCov_7']");
            }
           // Personal Property Replacement Cost (ML - 55)
               if (ExcelUtil.GetCellData(row,3, sheetname0)!="YES" && ExcelUtil.GetCellData(row, 3, sheetname0) != "")
               {
                   Perform.Click(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlHomSectionCoverages_Repeater1_ctlSectionCoverageItem_8_chkCov_8']");
               }
            //Backup of Sewer or Drain (92-173)
            if (ExcelUtil.GetCellData(row,4, sheetname0)!="YES" && ExcelUtil.GetCellData(row, 4, sheetname0) != "")
               {
                  // Perform.Click(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlHomSectionCoverages_Repeater1_ctlSectionCoverageItem_8_chkCov_8']");
                Perform.SelectDropDown(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlHomSectionCoverages_Repeater1_ctlSectionCoverageItem_9_ddIncreasedLimit_9']", ExcelUtil.GetCellData(row, 6, sheetname0));
            }
           
            //Cov.A - Specified Additional Amount Of Insurance (29-034) 
            if (ExcelUtil.GetCellData(row,8, sheetname0) == "YES" && ExcelUtil.GetCellData(row, 8, sheetname0) != "")
               {
                   Perform.Click(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlHomSectionCoverages_Repeater1_ctlSectionCoverageItem_10_chkCov_10']");
               }
            //Earthquake (HO-315B) 
            if (ExcelUtil.GetCellData(row,9, sheetname0) == "YES" && ExcelUtil.GetCellData(row, 9, sheetname0) != "")
               {
                   Perform.Click(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlHomSectionCoverages_Repeater1_ctlSectionCoverageItem_11_chkCov_11']");
                   Perform.SelectDropDown(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlHomSectionCoverages_Repeater1_ctlSectionCoverageItem_11_ddDeductible_11']", ExcelUtil.GetCellData(row, 10, sheetname0));
               }
            if (Perform.driver.FindElement(By.XPath(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlHomSectionCoverages_lnkMoreLess']")).Displayed)
            {

                Perform.Click(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlHomSectionCoverages_lnkMoreLess']");
            }
            //Earthquake (ML-54) 
            if (ExcelUtil.GetCellData(row,11,sheetname0)=="YES" && ExcelUtil.GetCellData(row, 11, sheetname0) != "")
               {
                   Perform.Click(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlHomSectionCoverages_Repeater1_ctlSectionCoverageItem_11_chkCov_11']");
                Perform.SelectDropDown(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlHomSectionCoverages_Repeater1_ctlSectionCoverageItem_11_ddDeductible_11']", ExcelUtil.GetCellData(row, 12, sheetname0));
               }
            //Actual Cash Value Loss Settlement (HO-04 81) 
            if (ExcelUtil.GetCellData(row, 13,sheetname0) == "YES" && ExcelUtil.GetCellData(row, 13, sheetname0) != "")
               {
                   Perform.Click(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlHomSectionCoverages_Repeater2_ctlSectionCoverageItem_0_chkCov_0']");
               }
            // Actual Cash Value Loss Settlement/Windstorm or Hail Losses to Roof Surfacing (HO-04 93) 
            if (ExcelUtil.GetCellData(row, 14, sheetname0) == "YES" && ExcelUtil.GetCellData(row, 14, sheetname0) != "")
               {
                Perform.Click(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlHomSectionCoverages_Repeater2_ctlSectionCoverageItem_1_chkCov_1']");
               }
            //Credit Card, Fund Transfer Card, Forgery and Counterfeit Money Coverage (HO-53
            if (ExcelUtil.GetCellData(row, 15,sheetname0)!="")
               {
                Perform.SelectDropDown(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlHomSectionCoverages_Repeater2_ctlSectionCoverageItem_3_ddIncreasedLimit_3']", ExcelUtil.GetCellData(row, 17, sheetname0));
               }
            //Loss Assessment (HO-35) 
            if (ExcelUtil.GetCellData(row, 19,sheetname0)!="")
               {
                   Perform.Click(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlHomSectionCoverages_Repeater2_ctlSectionCoverageItem_5_chkCov_5']");
                Perform.SelectDropDown(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlHomSectionCoverages_Repeater2_ctlSectionCoverageItem_5_ddIncreasedLimit_5']", ExcelUtil.GetCellData(row, 21, sheetname0));
               }
            //Loss Assessment - Earthquake (HO-35B) 
            if (ExcelUtil.GetCellData(row, 23,sheetname0)=="YES")
               {
                   Perform.Click(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlHomSectionCoverages_Repeater2_ctlSectionCoverageItem_6_chkCov_6']");
                Perform.EnterText(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlHomSectionCoverages_Repeater2_ctlSectionCoverageItem_6_txtLimit_6']", ExcelUtil.GetCellData(row, 24, sheetname0));
               }
               /*if (ExcelUtil.GetCellData(row, 24,sheetname0) == "NO")
               {
                   Perform.Click(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlHomSectionCoverages_Repeater2_ctlSectionCoverageItem_10_chkCov_10']");
               }*/
               Console.WriteLine("Optional Coverage is enetered");
            //Inland Marine
            Perform.waitTillElementToAppear(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlIMRVWatercraft_ctlInlandMarine_lblInlandMarineHdr']");
            Perform.Click(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlIMRVWatercraft_ctlInlandMarine_lblInlandMarineHdr']");
            //Jewelry
            string sheetname1 = "INLAND_MARINE";
      
            if  (ExcelUtil.GetCellData(row,0,sheetname1)=="YES")
            {
        
                Perform.Click(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlIMRVWatercraft_ctlInlandMarine_chkIMJewelry']");
                Perform.Click(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlIMRVWatercraft_ctlInlandMarine_lblInlandMarineHdr']");
                Perform.waitTillElementToAppear(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlIMRVWatercraft_ctlInlandMarine_ctlJewelry_IMRepeater_ctlInlandMarineIncreasedLimit_0_txtIM_LimitData_0']");
                Perform.EnterText(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlIMRVWatercraft_ctlInlandMarine_ctlJewelry_IMRepeater_ctlInlandMarineIncreasedLimit_0_txtIM_LimitData_0']", ExcelUtil.GetCellData(row, 1, sheetname1));
                Perform.SelectDropDown(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlIMRVWatercraft_ctlInlandMarine_ctlJewelry_IMRepeater_ctlInlandMarineIncreasedLimit_0_ddlIM_Deductible_0']",ExcelUtil.GetCellData(row,2,sheetname1));
                //Perform.waitTillElementToAppear(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlIMRVWatercraft_ctlInlandMarine_ctlJewelry_IMRepeater_ctlInlandMarineIncreasedLimit_0_txtIM_Description_0']");
                Perform.EnterText(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlIMRVWatercraft_ctlInlandMarine_ctlJewelry_IMRepeater_ctlInlandMarineIncreasedLimit_0_txtIM_Description_0']", ExcelUtil.GetCellData(row, 3, sheetname1));
                System.Threading.Thread.Sleep(500);
            }
            //Jewelry in Vault
            if (ExcelUtil.GetCellData(row,4, sheetname1) == "YES")
            {
                Perform.Click(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlIMRVWatercraft_ctlInlandMarine_chkJewelInVault']");
                Perform.waitTillElementToAppear(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlIMRVWatercraft_ctlInlandMarine_ctlJewelInVault_IMRepeater_ctlInlandMarineIncreasedLimit_0_txtIM_LimitData_0']");
                Perform.EnterText(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlIMRVWatercraft_ctlInlandMarine_ctlJewelInVault_IMRepeater_ctlInlandMarineIncreasedLimit_0_txtIM_LimitData_0']", ExcelUtil.GetCellData(row, 5, sheetname1));
                Perform.SelectDropDown(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlIMRVWatercraft_ctlInlandMarine_ctlJewelInVault_IMRepeater_ctlInlandMarineIncreasedLimit_0_ddlIM_Deductible_0']", ExcelUtil.GetCellData(row, 6, sheetname1));
                Perform.EnterText(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlIMRVWatercraft_ctlInlandMarine_ctlJewelInVault_IMRepeater_ctlInlandMarineIncreasedLimit_0_txtIM_Description_0']", ExcelUtil.GetCellData(row, 7, sheetname1));
            }
            //Bicycle
            if (ExcelUtil.GetCellData(row,8, sheetname1) == "YES")
            {
                Perform.waitTillElementToAppear(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlIMRVWatercraft_ctlInlandMarine_chkBike']");
                Perform.Click(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlIMRVWatercraft_ctlInlandMarine_chkBike']");
               // Perform.Click(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlIMRVWatercraft_ctlInlandMarine_lblInlandMarineHdr']");
                Perform.waitTillElementToAppear(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlIMRVWatercraft_ctlInlandMarine_ctlBikeList_IMRepeater_ctlInlandMarineIncreasedLimit_0_txtIM_LimitData_0']");
                Perform.EnterText(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlIMRVWatercraft_ctlInlandMarine_ctlBikeList_IMRepeater_ctlInlandMarineIncreasedLimit_0_txtIM_LimitData_0']", ExcelUtil.GetCellData(row, 9, sheetname1));
                Perform.SelectDropDown(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlIMRVWatercraft_ctlInlandMarine_ctlBikeList_IMRepeater_ctlInlandMarineIncreasedLimit_0_ddlIM_Deductible_0']", ExcelUtil.GetCellData(row, 10, sheetname1));
                Perform.EnterText(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlIMRVWatercraft_ctlInlandMarine_ctlBikeList_IMRepeater_ctlInlandMarineIncreasedLimit_0_txtIM_Description_0']", ExcelUtil.GetCellData(row, 11, sheetname1));

            }
            //Camera
            if (ExcelUtil.GetCellData(row,12, sheetname1) == "YES")
            {
                Perform.Click(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlIMRVWatercraft_ctlInlandMarine_chkCameras']");
                Perform.waitTillElementToAppear(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlIMRVWatercraft_ctlInlandMarine_ctlCameras_IMRepeater_ctlInlandMarineIncreasedLimit_0_txtIM_LimitData_0']");
                Perform.EnterText(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlIMRVWatercraft_ctlInlandMarine_ctlCameras_IMRepeater_ctlInlandMarineIncreasedLimit_0_txtIM_LimitData_0']", ExcelUtil.GetCellData(row, 13, sheetname1));
                Perform.SelectDropDown(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlIMRVWatercraft_ctlInlandMarine_ctlCameras_IMRepeater_ctlInlandMarineIncreasedLimit_0_ddlIM_Deductible_0']", ExcelUtil.GetCellData(row, 14, sheetname1));
                Perform.EnterText(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlIMRVWatercraft_ctlInlandMarine_ctlCameras_IMRepeater_ctlInlandMarineIncreasedLimit_0_txtIM_Description_0']", ExcelUtil.GetCellData(row, 15, sheetname1));
            }
            //Coins
            if (ExcelUtil.GetCellData(row,16, sheetname1) == "YES")
            {
                Perform.Click(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlIMRVWatercraft_ctlInlandMarine_chkCoins']");
                Perform.waitTillElementToAppear(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlIMRVWatercraft_ctlInlandMarine_CtlCoinsList_IMRepeater_ctlInlandMarineIncreasedLimit_0_txtIM_LimitData_0']");
                Perform.EnterText(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlIMRVWatercraft_ctlInlandMarine_CtlCoinsList_IMRepeater_ctlInlandMarineIncreasedLimit_0_txtIM_LimitData_0']", ExcelUtil.GetCellData(row, 17, sheetname1));
                Perform.SelectDropDown(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlIMRVWatercraft_ctlInlandMarine_CtlCoinsList_IMRepeater_ctlInlandMarineIncreasedLimit_0_ddlIM_Deductible_0']", ExcelUtil.GetCellData(row, 18, sheetname1));
                Perform.EnterText(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlIMRVWatercraft_ctlInlandMarine_CtlCoinsList_IMRepeater_ctlInlandMarineIncreasedLimit_0_txtIM_Description_0']", ExcelUtil.GetCellData(row, 19, sheetname1));

            }
            //Computers
            if (ExcelUtil.GetCellData(row,20, sheetname1) == "YES")
            {
                Perform.waitTillElementToAppear(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlIMRVWatercraft_ctlInlandMarine_chkComputers']");
                Perform.Click(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlIMRVWatercraft_ctlInlandMarine_chkComputers']");
                Perform.waitTillElementToAppear(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlIMRVWatercraft_ctlInlandMarine_ctlComputers_IMRepeater_ctlInlandMarineIncreasedLimit_0_txtIM_LimitData_0']");
                Perform.EnterText(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlIMRVWatercraft_ctlInlandMarine_ctlComputers_IMRepeater_ctlInlandMarineIncreasedLimit_0_txtIM_LimitData_0']", ExcelUtil.GetCellData(row, 21, sheetname1));
                Perform.SelectDropDown(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlIMRVWatercraft_ctlInlandMarine_ctlComputers_IMRepeater_ctlInlandMarineIncreasedLimit_0_ddlIM_Deductible_0']", ExcelUtil.GetCellData(row, 22, sheetname1));
                Perform.EnterText(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlIMRVWatercraft_ctlInlandMarine_ctlComputers_IMRepeater_ctlInlandMarineIncreasedLimit_0_txtIM_Description_0']", ExcelUtil.GetCellData(row, 23, sheetname1));
            }
            //Farm Machinery-Scheduled
            if (ExcelUtil.GetCellData(row,24, sheetname1) == "YES")
            {
                Perform.waitTillElementToAppear(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlIMRVWatercraft_ctlInlandMarine_chkFarmMachineSched']");
                Perform.Click(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlIMRVWatercraft_ctlInlandMarine_chkFarmMachineSched']");
                Perform.waitTillElementToAppear(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlIMRVWatercraft_ctlInlandMarine_ctlFarmMachineSched_IMRepeater_ctlInlandMarineIncreasedLimit_0_txtIM_LimitData_0']");
                Perform.EnterText(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlIMRVWatercraft_ctlInlandMarine_ctlFarmMachineSched_IMRepeater_ctlInlandMarineIncreasedLimit_0_txtIM_LimitData_0']", ExcelUtil.GetCellData(row, 25, sheetname1));
                Perform.SelectDropDown(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlIMRVWatercraft_ctlInlandMarine_ctlFarmMachineSched_IMRepeater_ctlInlandMarineIncreasedLimit_0_ddlIM_Deductible_0']", ExcelUtil.GetCellData(row, 26, sheetname1));
                Perform.EnterText(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlIMRVWatercraft_ctlInlandMarine_ctlFarmMachineSched_IMRepeater_ctlInlandMarineIncreasedLimit_0_txtIM_Description_0']", ExcelUtil.GetCellData(row, 27, sheetname1));
            }
            //Fine Arts-with breakage coverage
            if (ExcelUtil.GetCellData(row,28, sheetname1) == "YES")
            {
                Perform.Click(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlIMRVWatercraft_ctlInlandMarine_chkFABreak']");
                Perform.waitTillElementToAppear(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlIMRVWatercraft_ctlInlandMarine_ctlArtsBreak_IMRepeater_ctlInlandMarineIncreasedLimit_0_txtIM_LimitData_0']");
                Perform.EnterText(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlIMRVWatercraft_ctlInlandMarine_ctlArtsBreak_IMRepeater_ctlInlandMarineIncreasedLimit_0_txtIM_LimitData_0']", ExcelUtil.GetCellData(row, 29, sheetname1));
                Perform.SelectDropDown(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlIMRVWatercraft_ctlInlandMarine_ctlArtsBreak_IMRepeater_ctlInlandMarineIncreasedLimit_0_ddlIM_Deductible_0']", ExcelUtil.GetCellData(row, 30, sheetname1));
                Perform.EnterText(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlIMRVWatercraft_ctlInlandMarine_ctlArtsBreak_IMRepeater_ctlInlandMarineIncreasedLimit_0_txtIM_Description_0']", ExcelUtil.GetCellData(row, 31, sheetname1));
            }
            //Fine Arts-without breakeage coverage
            if (ExcelUtil.GetCellData(row,32, sheetname1) == "YES")
            {
                Perform.Click(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlIMRVWatercraft_ctlInlandMarine_chkFANoBreak']");
                Perform.waitTillElementToAppear(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlIMRVWatercraft_ctlInlandMarine_ctlArtsNoBreak_IMRepeater_ctlInlandMarineIncreasedLimit_0_txtIM_LimitData_0']");
                Perform.EnterText(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlIMRVWatercraft_ctlInlandMarine_ctlArtsNoBreak_IMRepeater_ctlInlandMarineIncreasedLimit_0_txtIM_LimitData_0']", ExcelUtil.GetCellData(row, 33, sheetname1));
                Perform.SelectDropDown(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlIMRVWatercraft_ctlInlandMarine_ctlArtsNoBreak_IMRepeater_ctlInlandMarineIncreasedLimit_0_ddlIM_Deductible_0']", ExcelUtil.GetCellData(row, 34, sheetname1));
                Perform.EnterText(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlIMRVWatercraft_ctlInlandMarine_ctlArtsNoBreak_IMRepeater_ctlInlandMarineIncreasedLimit_0_txtIM_Description_0']", ExcelUtil.GetCellData(row, 35, sheetname1));
            }
            //Fur
            if (ExcelUtil.GetCellData(row,36, sheetname1) == "YES")
            {
                Perform.Click(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlIMRVWatercraft_ctlInlandMarine_chkFurs']");
                Perform.waitTillElementToAppear(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlIMRVWatercraft_ctlInlandMarine_ctlFurs_IMRepeater_ctlInlandMarineIncreasedLimit_0_txtIM_LimitData_0']");
                Perform.EnterText(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlIMRVWatercraft_ctlInlandMarine_ctlFurs_IMRepeater_ctlInlandMarineIncreasedLimit_0_txtIM_LimitData_0']", ExcelUtil.GetCellData(row, 37, sheetname1));
                Perform.SelectDropDown(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlIMRVWatercraft_ctlInlandMarine_ctlFurs_IMRepeater_ctlInlandMarineIncreasedLimit_0_ddlIM_Deductible_0']", ExcelUtil.GetCellData(row, 38, sheetname1));
                Perform.EnterText(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlIMRVWatercraft_ctlInlandMarine_ctlFurs_IMRepeater_ctlInlandMarineIncreasedLimit_0_txtIM_Description_0']", ExcelUtil.GetCellData(row, 39, sheetname1));
            }
            //Garden Tractor
            if (ExcelUtil.GetCellData(row,40, sheetname1) == "YES")
            {
                Perform.Click(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlIMRVWatercraft_ctlInlandMarine_chkGarden']");
                Perform.waitTillElementToAppear(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlIMRVWatercraft_ctlInlandMarine_ctlGarden_IMRepeater_ctlInlandMarineIncreasedLimit_0_txtIM_LimitData_0']");
                Perform.EnterText(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlIMRVWatercraft_ctlInlandMarine_ctlGarden_IMRepeater_ctlInlandMarineIncreasedLimit_0_txtIM_LimitData_0']", ExcelUtil.GetCellData(row, 41, sheetname1));
                Perform.SelectDropDown(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlIMRVWatercraft_ctlInlandMarine_ctlGarden_IMRepeater_ctlInlandMarineIncreasedLimit_0_ddlIM_Deductible_0']", ExcelUtil.GetCellData(row, 42, sheetname1));
                Perform.EnterText(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlIMRVWatercraft_ctlInlandMarine_ctlGarden_IMRepeater_ctlInlandMarineIncreasedLimit_0_txtIM_Description_0']", ExcelUtil.GetCellData(row, 43, sheetname1));
            }
            //Golfers Equipment
            if (ExcelUtil.GetCellData(row,44, sheetname1) == "YES")
            {
                Perform.Click(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlIMRVWatercraft_ctlInlandMarine_chkGolfers']");
                Perform.waitTillElementToAppear(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlIMRVWatercraft_ctlInlandMarine_ctlGolfers_IMRepeater_ctlInlandMarineIncreasedLimit_0_txtIM_LimitData_0']");
                Perform.EnterText(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlIMRVWatercraft_ctlInlandMarine_ctlGolfers_IMRepeater_ctlInlandMarineIncreasedLimit_0_txtIM_LimitData_0']", ExcelUtil.GetCellData(row, 45, sheetname1));
                Perform.SelectDropDown(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlIMRVWatercraft_ctlInlandMarine_ctlGolfers_IMRepeater_ctlInlandMarineIncreasedLimit_0_ddlIM_Deductible_0']", ExcelUtil.GetCellData(row, 46, sheetname1));
                Perform.EnterText(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlIMRVWatercraft_ctlInlandMarine_ctlGolfers_IMRepeater_ctlInlandMarineIncreasedLimit_0_txtIM_Description_0']", ExcelUtil.GetCellData(row, 47, sheetname1));
            }
            //Guns
            if (ExcelUtil.GetCellData(row,48, sheetname1) == "YES")
            {
                Perform.Click(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlIMRVWatercraft_ctlInlandMarine_chkGuns']");
                Perform.waitTillElementToAppear(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlIMRVWatercraft_ctlInlandMarine_ctlGuns_IMRepeater_ctlInlandMarineIncreasedLimit_0_txtIM_LimitData_0']");
                Perform.EnterText(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlIMRVWatercraft_ctlInlandMarine_ctlGuns_IMRepeater_ctlInlandMarineIncreasedLimit_0_txtIM_LimitData_0']", ExcelUtil.GetCellData(row, 49, sheetname1));
                Perform.SelectDropDown(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlIMRVWatercraft_ctlInlandMarine_ctlGuns_IMRepeater_ctlInlandMarineIncreasedLimit_0_ddlIM_Deductible_0']", ExcelUtil.GetCellData(row, 50, sheetname1));
                Perform.EnterText(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlIMRVWatercraft_ctlInlandMarine_ctlGuns_IMRepeater_ctlInlandMarineIncreasedLimit_0_txtIM_Description_0']", ExcelUtil.GetCellData(row, 51, sheetname1));

            }
            //Hearing Aid
            if (ExcelUtil.GetCellData(row,52, sheetname1) == "YES")
            {
                System.Threading.Thread.Sleep(500);
                Perform.Click(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlIMRVWatercraft_ctlInlandMarine_chkHearing']");
                Perform.waitTillElementToAppear(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlIMRVWatercraft_ctlInlandMarine_ctlHearing_IMRepeater_ctlInlandMarineIncreasedLimit_0_txtIM_LimitData_0']");
                Perform.EnterText(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlIMRVWatercraft_ctlInlandMarine_ctlHearing_IMRepeater_ctlInlandMarineIncreasedLimit_0_txtIM_LimitData_0']", ExcelUtil.GetCellData(row, 53, sheetname1));
                Perform.SelectDropDown(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlIMRVWatercraft_ctlInlandMarine_ctlHearing_IMRepeater_ctlInlandMarineIncreasedLimit_0_ddlIM_Deductible_0']", ExcelUtil.GetCellData(row, 54, sheetname1));
                Perform.EnterText(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlIMRVWatercraft_ctlInlandMarine_ctlHearing_IMRepeater_ctlInlandMarineIncreasedLimit_0_txtIM_Description_0']", ExcelUtil.GetCellData(row, 55, sheetname1));

            }
            //Musical Instrument
            if (ExcelUtil.GetCellData(row,56, sheetname1) == "YES")
            {
                Perform.Click(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlIMRVWatercraft_ctlInlandMarine_chkMusic']");
                Perform.waitTillElementToAppear(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlIMRVWatercraft_ctlInlandMarine_ctlMusicInstr_IMRepeater_ctlInlandMarineIncreasedLimit_0_txtIM_LimitData_0']");
                Perform.EnterText(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlIMRVWatercraft_ctlInlandMarine_ctlMusicInstr_IMRepeater_ctlInlandMarineIncreasedLimit_0_txtIM_LimitData_0']", ExcelUtil.GetCellData(row, 57, sheetname1));
                Perform.SelectDropDown(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlIMRVWatercraft_ctlInlandMarine_ctlMusicInstr_IMRepeater_ctlInlandMarineIncreasedLimit_0_ddlIM_Deductible_0']", ExcelUtil.GetCellData(row, 58, sheetname1));
                Perform.EnterText(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlIMRVWatercraft_ctlInlandMarine_ctlMusicInstr_IMRepeater_ctlInlandMarineIncreasedLimit_0_txtIM_Description_0']", ExcelUtil.GetCellData(row, 59, sheetname1));

            }
            //Silverware
            if (ExcelUtil.GetCellData(row,60, sheetname1) == "YES")
            {
                Perform.Click(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlIMRVWatercraft_ctlInlandMarine_chkSilverware']");
                Perform.waitTillElementToAppear(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlIMRVWatercraft_ctlInlandMarine_ctlSilverware_IMRepeater_ctlInlandMarineIncreasedLimit_0_txtIM_LimitData_0']");
                Perform.EnterText(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlIMRVWatercraft_ctlInlandMarine_ctlSilverware_IMRepeater_ctlInlandMarineIncreasedLimit_0_txtIM_LimitData_0']", ExcelUtil.GetCellData(row, 61, sheetname1));
                Perform.SelectDropDown(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlIMRVWatercraft_ctlInlandMarine_ctlSilverware_IMRepeater_ctlInlandMarineIncreasedLimit_0_ddlIM_Deductible_0']", ExcelUtil.GetCellData(row, 62, sheetname1));
                Perform.EnterText(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlIMRVWatercraft_ctlInlandMarine_ctlSilverware_IMRepeater_ctlInlandMarineIncreasedLimit_0_txtIM_Description_0']", ExcelUtil.GetCellData(row, 63, sheetname1));
            }
            //Telephone
            if (ExcelUtil.GetCellData(row,64, sheetname1) == "YES")
            {
                Perform.Click(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlIMRVWatercraft_ctlInlandMarine_chkMobile']");
                Perform.waitTillElementToAppear(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlIMRVWatercraft_ctlInlandMarine_ctlMobile_IMRepeater_ctlInlandMarineIncreasedLimit_0_txtIM_LimitData_0']");
                Perform.EnterText(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlIMRVWatercraft_ctlInlandMarine_ctlMobile_IMRepeater_ctlInlandMarineIncreasedLimit_0_txtIM_LimitData_0']", ExcelUtil.GetCellData(row, 65, sheetname1));
                Perform.SelectDropDown(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlIMRVWatercraft_ctlInlandMarine_ctlMobile_IMRepeater_ctlInlandMarineIncreasedLimit_0_ddlIM_Deductible_0']", ExcelUtil.GetCellData(row, 66, sheetname1));
                Perform.EnterText(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlIMRVWatercraft_ctlInlandMarine_ctlMobile_IMRepeater_ctlInlandMarineIncreasedLimit_0_txtIM_Description_0']", ExcelUtil.GetCellData(row, 67, sheetname1));
            }
            //Tools and Equipments
            if (ExcelUtil.GetCellData(row,68, sheetname1) == "YES")
            {
                Perform.Click(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlIMRVWatercraft_ctlInlandMarine_chkTools']");
                Perform.waitTillElementToAppear("//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlIMRVWatercraft_ctlInlandMarine_ctlTools_IMRepeater_ctlInlandMarineIncreasedLimit_0_txtIM_LimitData_0']");
                Perform.EnterText(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlIMRVWatercraft_ctlInlandMarine_ctlTools_IMRepeater_ctlInlandMarineIncreasedLimit_0_txtIM_LimitData_0']", ExcelUtil.GetCellData(row, 69, sheetname1));
                Perform.SelectDropDown(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlIMRVWatercraft_ctlInlandMarine_ctlTools_IMRepeater_ctlInlandMarineIncreasedLimit_0_ddlIM_Deductible_0']", ExcelUtil.GetCellData(row, 70, sheetname1));
                Perform.EnterText(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlIMRVWatercraft_ctlInlandMarine_ctlTools_IMRepeater_ctlInlandMarineIncreasedLimit_0_txtIM_Description_0']", ExcelUtil.GetCellData(row, 71, sheetname1));

            }
            Console.WriteLine("Inland Marine details are entered");
            //Add RV/WaterCraft
            string sheetname2 = "RV_WATERCRAFT";
            if (ExcelUtil.GetCellData(row, 0, sheetname2) != "")
            {
                System.Threading.Thread.Sleep(500);
                    Perform.driver.FindElement(By.XPath(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlIMRVWatercraft_lnkAddRVWater']")).SendKeys(Keys.PageDown);
                    Perform.Click(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlIMRVWatercraft_lnkAddRVWater']");

                
                Perform.waitTillElementToAppear(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlIMRVWatercraft_ctlRV_WatercraftList_rvWaterRepeater_ctlRV_Watercraft_0_ddlVehType_0']");
                if (ExcelUtil.GetCellData(row, 0, sheetname2) == "WATERCRAFT")
                {

                    Perform.SelectDropDown(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlIMRVWatercraft_ctlRV_WatercraftList_rvWaterRepeater_ctlRV_Watercraft_0_ddlVehType_0']", ExcelUtil.GetCellData(row, 0, sheetname2));
                    Perform.driver.SwitchTo().Alert().Accept();
                    Perform.SelectDropDown(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlIMRVWatercraft_ctlRV_WatercraftList_rvWaterRepeater_ctlRV_Watercraft_0_ddlCoverageOptions_0']", ExcelUtil.GetCellData(row, 1, sheetname2));
                    Perform.SelectDropDown(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlIMRVWatercraft_ctlRV_WatercraftList_rvWaterRepeater_ctlRV_Watercraft_0_ddlBodilyInjuryLimit_0']", ExcelUtil.GetCellData(row, 2, sheetname2));
                    if (ExcelUtil.GetCellData(row, 3, sheetname2) == "YES")
                    {
                        Perform.Click(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlIMRVWatercraft_ctlRV_WatercraftList_rvWaterRepeater_ctlRV_Watercraft_0_chkUnder25Operator_0']");
                    }
                    Perform.EnterText(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlIMRVWatercraft_ctlRV_WatercraftList_rvWaterRepeater_ctlRV_Watercraft_0_txtVehYear_0']", ExcelUtil.GetCellData(row, 4, sheetname2));
                    Perform.EnterText(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlIMRVWatercraft_ctlRV_WatercraftList_rvWaterRepeater_ctlRV_Watercraft_0_txtVehLength_0']", ExcelUtil.GetCellData(row, 5, sheetname2));
                    Perform.SelectDropDown(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlIMRVWatercraft_ctlRV_WatercraftList_rvWaterRepeater_ctlRV_Watercraft_0_ddlMotorType_0']", ExcelUtil.GetCellData(row, 6, sheetname2));
                    Perform.EnterText(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlIMRVWatercraft_ctlRV_WatercraftList_rvWaterRepeater_ctlRV_Watercraft_0_txtVehSerialNum_0']", ExcelUtil.GetCellData(row, 7, sheetname2));
                    Perform.EnterText(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlIMRVWatercraft_ctlRV_WatercraftList_rvWaterRepeater_ctlRV_Watercraft_0_txtVehMake_0']", ExcelUtil.GetCellData(row, 8, sheetname2));
                    Perform.EnterText(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlIMRVWatercraft_ctlRV_WatercraftList_rvWaterRepeater_ctlRV_Watercraft_0_txtVehModel_0']", ExcelUtil.GetCellData(row, 9, sheetname2));
                    if (ExcelUtil.GetCellData(row, 10, sheetname2) != "")
                    {
                        Perform.SelectDropDown(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlIMRVWatercraft_ctlRV_WatercraftList_rvWaterRepeater_ctlRV_Watercraft_0_ddlPropertyDeductible_0']", ExcelUtil.GetCellData(row, 10, sheetname2));
                    }
                    if (ExcelUtil.GetCellData(row, 11, sheetname2) != "")
                    {
                        Perform.EnterText(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlIMRVWatercraft_ctlRV_WatercraftList_rvWaterRepeater_ctlRV_Watercraft_0_txtVehCostNew_0']", ExcelUtil.GetCellData(row, 11, sheetname2));
                    }
                }

                if (ExcelUtil.GetCellData(row, 0, sheetname2) == "BOAT MOTOR ONLY")
                {
                    Perform.SelectDropDown(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlIMRVWatercraft_ctlRV_WatercraftList_rvWaterRepeater_ctlRV_Watercraft_0_ddlVehType_0']", ExcelUtil.GetCellData(row, 0, sheetname2));
                    Perform.driver.SwitchTo().Alert().Accept();
                    Perform.SelectDropDown(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlIMRVWatercraft_ctlRV_WatercraftList_rvWaterRepeater_ctlRV_Watercraft_0_ddlPropertyDeductible_0']", ExcelUtil.GetCellData(row, 10, sheetname2));

                    Perform.EnterText(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlIMRVWatercraft_ctlRV_WatercraftList_rvWaterRepeater_ctlRV_Watercraft_0_txtMotorCostNew_0']", ExcelUtil.GetCellData(row, 11, sheetname2));
                    Perform.EnterText(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlIMRVWatercraft_ctlRV_WatercraftList_rvWaterRepeater_ctlRV_Watercraft_0_txtHorsepowerCCs_0']", ExcelUtil.GetCellData(row, 12, sheetname2));
                    Perform.EnterText(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlIMRVWatercraft_ctlRV_WatercraftList_rvWaterRepeater_ctlRV_Watercraft_0_txtMotorYear_0']", ExcelUtil.GetCellData(row, 4, sheetname2));
                    Perform.EnterText(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlIMRVWatercraft_ctlRV_WatercraftList_rvWaterRepeater_ctlRV_Watercraft_0_txtMotorSerialNum_0']", ExcelUtil.GetCellData(row, 7, sheetname2));
                    Perform.EnterText(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlIMRVWatercraft_ctlRV_WatercraftList_rvWaterRepeater_ctlRV_Watercraft_0_txtMotorMake_0']", ExcelUtil.GetCellData(row, 8, sheetname2));
                    Perform.EnterText(".//*[@id='cphMain_ctlHomeInput_ctlCoverages_HOM_ctlIMRVWatercraft_ctlRV_WatercraftList_rvWaterRepeater_ctlRV_Watercraft_0_txtMotorModel_0']", ExcelUtil.GetCellData(row, 9, sheetname2));

                }
            }
            else
                Console.WriteLine("No RV/Watercraft is present");
            
        }
        public static void Application(int row, string sheetname)
        {
            sheetname = "Application";
            if (ExcelUtil.GetCellData(row, 0, sheetname) != "")
            {
                Perform.EnterText(".//*[@id='cphMain_ctl_Master_HOM_APP_ctl_HOM_App_Section_ctl_PropertyUpdates_HOM_App_txtMobileHomeLength']", ExcelUtil.GetCellData(row, 0, sheetname));
                Perform.EnterText(".//*[@id='cphMain_ctl_Master_HOM_APP_ctl_HOM_App_Section_ctl_PropertyUpdates_HOM_App_txtMobileHomeWidth']", ExcelUtil.GetCellData(row, 1, sheetname));
            }
            if (ExcelUtil.GetCellData(row, 2, sheetname) != "")
            {
                if (Perform.driver.FindElement(By.XPath(".//*[@id='cphMain_ctl_Master_HOM_APP_ctl_HOM_App_Section_ctl_PropertyUpdates_HOM_App_txtRoofUpdateYear']")).Displayed)
                {
                    if (ExcelUtil.GetCellData(row, 3, sheetname) != "" && ExcelUtil.GetCellData(row, 3, sheetname) == "COMPLETE")
                    {
                       // Perform.waitTillElementToAppear("//label[contains(@for,'radRoofComplete')]");
                        // Perform.Click(".//*[@id='cphMain_ctl_Master_HOM_APP_ctl_HOM_App_Section_ctl_PropertyUpdates_HOM_App_radRoofComplete']");
                        Perform.Click("//label[contains(@for,'radRoofComplete')]");


                    }
                    else
                    {
                        Perform.Click(".//*[@id='cphMain_ctl_Master_HOM_APP_ctl_HOM_App_Section_ctl_PropertyUpdates_HOM_App_radRoofPartial']");
                      
                    }
                    Perform.SelectDropDown(".//*[@id='cphMain_ctl_Master_HOM_APP_ctl_HOM_App_Section_ctl_PropertyUpdates_HOM_App_ddRoofType']", ExcelUtil.GetCellData(row, 4, sheetname));
                    Perform.EnterText(".//*[@id='cphMain_ctl_Master_HOM_APP_ctl_HOM_App_Section_ctl_PropertyUpdates_HOM_App_txtRoofUpdateYear']", ExcelUtil.GetCellData(row, 2, sheetname));
                }
                else
                {
                    Perform.Click(".//*[@id='cphMain_ctl_Master_HOM_APP_ctl_HOM_App_Section_ctl_PropertyUpdates_HOM_App_lblMainAccord']");
                    Perform.waitTillElementToAppear(".//*[@id='cphMain_ctl_Master_HOM_APP_ctl_HOM_App_Section_ctl_PropertyUpdates_HOM_App_txtRoofUpdateYear']");
                    if (ExcelUtil.GetCellData(row, 3, sheetname) != "" && ExcelUtil.GetCellData(row, 3, sheetname) == "COMPLETE")
                    {
                        //Perform.Click(".//*[@id='cphMain_ctl_Master_HOM_APP_ctl_HOM_App_Section_ctl_PropertyUpdates_HOM_App_radRoofComplete']");
                        Perform.Click("//label[contains(@for,'radRoofComplete')]");
                    }
                    else
                    {
                        Perform.Click(".//*[@id='cphMain_ctl_Master_HOM_APP_ctl_HOM_App_Section_ctl_PropertyUpdates_HOM_App_radRoofPartial']");
                    }
                    Perform.SelectDropDown(".//*[@id='cphMain_ctl_Master_HOM_APP_ctl_HOM_App_Section_ctl_PropertyUpdates_HOM_App_ddRoofType']", ExcelUtil.GetCellData(row, 4, sheetname));
                    Perform.EnterText(".//*[@id='cphMain_ctl_Master_HOM_APP_ctl_HOM_App_Section_ctl_PropertyUpdates_HOM_App_txtRoofUpdateYear']", ExcelUtil.GetCellData(row, 2, sheetname));
                }
                if (ExcelUtil.GetCellData(row, 6, sheetname) != "" && ExcelUtil.GetCellData(row, 6, sheetname) == "PARTIAL")
                {
                    Perform.Click(".//*[@id='cphMain_ctl_Master_HOM_APP_ctl_HOM_App_Section_ctl_PropertyUpdates_HOM_App_radCentralPartial']");
                }
                else
                {
                    Perform.Click(".//*[@id='cphMain_ctl_Master_HOM_APP_ctl_HOM_App_Section_ctl_PropertyUpdates_HOM_App_radCentralComplete']");
                }
                Perform.SelectDropDown(".//*[@id='cphMain_ctl_Master_HOM_APP_ctl_HOM_App_Section_ctl_PropertyUpdates_HOM_App_ddCentralAirType']", ExcelUtil.GetCellData(row, 7, sheetname));
                Perform.EnterText(".//*[@id='cphMain_ctl_Master_HOM_APP_ctl_HOM_App_Section_ctl_PropertyUpdates_HOM_App_txtCentralAirUpdated']", ExcelUtil.GetCellData(row, 5, sheetname));
                if (ExcelUtil.GetCellData(row, 9, sheetname) != "" && ExcelUtil.GetCellData(row, 9, sheetname) == "PARTIAL")
                {
                    Perform.Click(".//*[@id='cphMain_ctl_Master_HOM_APP_ctl_HOM_App_Section_ctl_PropertyUpdates_HOM_App_radElectricPartial']");
                }
                else
                {
                    Perform.Click(".//*[@id='cphMain_ctl_Master_HOM_APP_ctl_HOM_App_Section_ctl_PropertyUpdates_HOM_App_radElectricComplete']");
                }
                Perform.SelectDropDown(".//*[@id='cphMain_ctl_Master_HOM_APP_ctl_HOM_App_Section_ctl_PropertyUpdates_HOM_App_ddElectricType']", ExcelUtil.GetCellData(row, 10, sheetname));
                Perform.EnterText(".//*[@id='cphMain_ctl_Master_HOM_APP_ctl_HOM_App_Section_ctl_PropertyUpdates_HOM_App_txtElectricUpdated']", ExcelUtil.GetCellData(row, 8, sheetname));
                if (ExcelUtil.GetCellData(row, 12, sheetname) != "" && ExcelUtil.GetCellData(row, 12, sheetname) == "PARTIAL")
                {
                    Perform.Click(".//*[@id='cphMain_ctl_Master_HOM_APP_ctl_HOM_App_Section_ctl_PropertyUpdates_HOM_App_radPlumbingPartial']");
                }
                else
                {
                    Perform.Click(".//*[@id='cphMain_ctl_Master_HOM_APP_ctl_HOM_App_Section_ctl_PropertyUpdates_HOM_App_radPlumbingComplete']");
                }
                Perform.SelectDropDown(".//*[@id='cphMain_ctl_Master_HOM_APP_ctl_HOM_App_Section_ctl_PropertyUpdates_HOM_App_ddPlumbingType']", ExcelUtil.GetCellData(row, 13, sheetname));
                Perform.EnterText(".//*[@id='cphMain_ctl_Master_HOM_APP_ctl_HOM_App_Section_ctl_PropertyUpdates_HOM_App_txtPlumbingUpdated']", ExcelUtil.GetCellData(row, 11, sheetname));
                if (ExcelUtil.GetCellData(row, 15, sheetname) != "" && ExcelUtil.GetCellData(row, 15, sheetname) == "PARTIAL")
                {
                    Perform.Click(".//*[@id='cphMain_ctl_Master_HOM_APP_ctl_HOM_App_Section_ctl_PropertyUpdates_HOM_App_radInspectionPartial']");
                }
                else
                {
                    Perform.Click(".//*[@id='cphMain_ctl_Master_HOM_APP_ctl_HOM_App_Section_ctl_PropertyUpdates_HOM_App_radInspectionComplete']");
                }
                Perform.EnterText(".//*[@id='cphMain_ctl_Master_HOM_APP_ctl_HOM_App_Section_ctl_PropertyUpdates_HOM_App_txtInspectionDate']", ExcelUtil.GetCellData(row, 14, sheetname));
                Perform.EnterText(".//*[@id='cphMain_ctl_Master_HOM_APP_ctl_HOM_App_Section_ctl_PropertyUpdates_HOM_App_txtInspectionRemarks']", ExcelUtil.GetCellData(row, 16, sheetname));
            }
        }
    }
}
