using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using System;
using System.Data;

namespace VR_COM_BOP
{
    class PageUtility
    {
        public static void LocationCoverage(int row, string noflocation, string sheetname)
        {
            int col = 0;

            //sheetname = "Location_0";
            for (int i = 0; i < Int32.Parse(noflocation); i++)
            {
                int buildcol = 19;
                sheetname = "Location_" + i;
                if (i > 0)
                {
                    /* IWebElement but=Perform.driver.FindElement(By.XPath("//div[@id='divEditControls']/div/div[3]/input"));


                     Actions builder = new Actions(Perform.driver);
                     builder.MoveToElement(but,122,28).Click().Build().Perform();*/
                    Perform.driver.FindElement(By.XPath("//input[contains(@id,'btnAddAnotherLocation')]")).SendKeys(Keys.PageDown);

                    Perform.Click("//input[contains(@id,'btnAddAnotherLocation')]", Property_type.XPath);
                    Perform.driver.Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(10);
                    // Perform.waitTillElementToAppear("//div[3]/input");
                    // Perform.Click("//div[@id='divEditControls']/div/div[3]/input", Property_type.XPath);
                    Console.WriteLine("Add another location is clicked");
                }
                //if (ExcelUtil.GetCellData(row, col, sheetname) != "")
                //{ 

                //Perform.Click(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_LocationList_Repeater1_ctl_BOP_Location_"+i+"_ctlProperty_Address_"+i+"_btnCopyAddress_"+i+"']", Property_type.XPath);
                Perform.Wait();
                Perform.EnterText(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_LocationList_Repeater1_ctl_BOP_Location_" + i + "_ctlProperty_Address_" + i + "_txtStreetNum_" + i + "']", ExcelUtil.GetCellData(row, col, sheetname), Property_type.XPath);

                Perform.EnterText(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_LocationList_Repeater1_ctl_BOP_Location_" + i + "_ctlProperty_Address_" + i + "_txtStreetName_" + i + "']", ExcelUtil.GetCellData(row, col + 1, sheetname), Property_type.XPath);
                Perform.EnterText(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_LocationList_Repeater1_ctl_BOP_Location_" + i + "_ctlProperty_Address_" + i + "_txtNumberOfAmusementAreas_" + i + "']", ExcelUtil.GetCellData(row, col + 2, sheetname), Property_type.XPath);
                Perform.EnterText(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_LocationList_Repeater1_ctl_BOP_Location_" + i + "_ctlProperty_Address_" + i + "_txtNumberOfPlaygrounds_" + i + "']", ExcelUtil.GetCellData(row, col + 3, sheetname), Property_type.XPath);
                Perform.EnterText(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_LocationList_Repeater1_ctl_BOP_Location_" + i + "_ctlProperty_Address_" + i + "_txtNumberOfSwimmingPools_" + i + "']", ExcelUtil.GetCellData(row, col + 4, sheetname), Property_type.XPath);
                Perform.EnterText(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_LocationList_Repeater1_ctl_BOP_Location_" + i + "_ctlProperty_Address_" + i + "_txtZipCode_" + i + "']", ExcelUtil.GetCellData(row, col + 5, sheetname), Property_type.XPath);
                if (ExcelUtil.GetCellData(row, col + 6, sheetname) == "INDIANAPOLIS" || ExcelUtil.GetCellData(row, col + 6, sheetname) == "CHESTERTON" || ExcelUtil.GetCellData(row, col + 6, sheetname) == "TELL CITY" || ExcelUtil.GetCellData(row, col + 6, sheetname) == "JASPER" || ExcelUtil.GetCellData(row, col + 6, sheetname) == "NEWBURGH")
                {
                    Perform.SelectDropDown(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_LocationList_Repeater1_ctl_BOP_Location_"+i+"_ctlProperty_Address_"+i+"_ddCityName_"+i+"']", ExcelUtil.GetCellData(row, col + 7, sheetname), Property_type.XPath);
                    if(ExcelUtil.GetCellData(row,col+7,sheetname)=="--OTHER--")
                    {
                        Perform.driver.FindElement(By.XPath(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_LocationList_Repeater1_ctl_BOP_Location_"+i+"_ctlProperty_Address_"+i+"_txtCityName_"+i+"']")).Clear();
                        Perform.EnterText(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_LocationList_Repeater1_ctl_BOP_Location_"+i+"_ctlProperty_Address_"+i+"_txtCityName_"+i+"']", ExcelUtil.GetCellData(row, col + 8, sheetname), Property_type.XPath);
                    }
                    Console.WriteLine("City entered");
                }else
                {
                    Perform.driver.FindElement(By.XPath(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_LocationList_Repeater1_ctl_BOP_Location_" + i + "_ctlProperty_Address_" + i + "_txtCityName_" + i + "']")).Clear();
                    Perform.EnterText(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_LocationList_Repeater1_ctl_BOP_Location_" + i + "_ctlProperty_Address_" + i + "_txtCityName_" + i + "']", ExcelUtil.GetCellData(row, col + 6, sheetname), Property_type.XPath);
                }
                    Perform.SelectDropDown(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_LocationList_Repeater1_ctl_BOP_Location_"+ i +"_ctlProperty_Address_"+ i +"_ddStateAbbrev_"+i+"']", ExcelUtil.GetCellData(row, col + 9, sheetname), Property_type.XPath);
                Perform.driver.FindElement(By.XPath(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_LocationList_Repeater1_ctl_BOP_Location_" + i + "_ctlProperty_Address_" + i + "_txtGaragedCounty_" + i + "']")).Clear();
                    Perform.EnterText(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_LocationList_Repeater1_ctl_BOP_Location_" + i + "_ctlProperty_Address_" + i + "_txtGaragedCounty_" + i + "']", ExcelUtil.GetCellData(row, col + 10, sheetname), Property_type.XPath);
                    Console.WriteLine("Location" +( i + 1) + " details entered");

                

                    //Check for Optional Location Coverage Information

                    if (ExcelUtil.GetCellData(row, col + 10, sheetname) == "YES" && ExcelUtil.GetCellData(row, col + 9, sheetname)!="")
                {
                        Perform.Click(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_LocationList_Repeater1_ctl_BOP_Location_" + i + "_ctl_BOP_Location_Coverages_" + i + "_chkEquipmentBreakdown_" + i + "']", Property_type.XPath);
                        Perform.SelectDropDown(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_LocationList_Repeater1_ctl_BOP_Location_" + i + "_ctl_BOP_Location_Coverages_" + i + "_ddlEquipmentBreakdownDeductible_" + i + "']", ExcelUtil.GetCellData(row, col + 11, sheetname), Property_type.XPath);
                    }
                    //else
                    //{
                    // Console.WriteLine("No optional coverage information");
                    //}
                    if (ExcelUtil.GetCellData(row, col + 12, sheetname) == "YES" && ExcelUtil.GetCellData(row, col + 11, sheetname)!="")
                    {
                        Perform.Click(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_LocationList_Repeater1_ctl_BOP_Location_" + i + "_ctl_BOP_Location_Coverages_" + i + "_chkMoneySecuritiesONPremises_" + i + "']", Property_type.XPath);
                        Perform.EnterText(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_LocationList_Repeater1_ctl_BOP_Location_" + i + "_ctl_BOP_Location_Coverages_" + i + "_txtMoneySecuritiesONPremisesLimit_" + i + "']", ExcelUtil.GetCellData(row, col + 13, sheetname), Property_type.XPath);
                    }
                    if (ExcelUtil.GetCellData(row, col + 13, sheetname) == "YES" && ExcelUtil.GetCellData(row, col + 14, sheetname)!="")
                    {
                        Perform.Click(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_LocationList_Repeater1_ctl_BOP_Location_" + i + "_ctl_BOP_Location_Coverages_" + i + "_chkMoneySecuritiesOFFPremises_" + i + "']", Property_type.XPath);
                        Perform.EnterText(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_LocationList_Repeater1_ctl_BOP_Location_" + i + "_ctl_BOP_Location_Coverages_" + i + "_txtMoneySecuritiesOFFPremisesLimit_" + i + "']", ExcelUtil.GetCellData(row, col + 15, sheetname), Property_type.XPath);
                    }
                    if (ExcelUtil.GetCellData(row, col + 16, sheetname) == "YES" && ExcelUtil.GetCellData(row, col + 16, sheetname)!="")
                    {
                        Perform.Click(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_LocationList_Repeater1_ctl_BOP_Location_" + i + "_ctl_BOP_Location_Coverages_" + i + "_chkOutdoorSigns_" + i + "']", Property_type.XPath);
                        Perform.EnterText(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_LocationList_Repeater1_ctl_BOP_Location_" + i + "_ctl_BOP_Location_Coverages_" + i + "_txtOutdoorSignsLimit_" + i + "']", ExcelUtil.GetCellData(row, col + 17, sheetname), Property_type.XPath);
                    }
                

                        //Check for building information
                        String nofbuilding = ExcelUtil.GetCellData(row, 4 + i, "TestCase");
                for (int j = 0; j < Int32.Parse(nofbuilding); j++)
                {

                    // if (ExcelUtil.GetCellData(row, buildcol + 1, sheetname) != "")
                    //{
                    Perform.Click(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_LocationList_Repeater1_ctl_BOP_Location_" + i + "_btnAddBuilding_" + i + "']", Property_type.XPath);
                    Perform.driver.Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(10);
                    Perform.EnterText(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_LocationList_Repeater1_ctl_BOP_Location_" + i + "_ctl_BOP_BuildingList_" + i + "_Repeater1_" + i + "_ctl_BOP_Building_" + j + "_ctl_BOP_Building_Information_" + j + "_txtDescription_" + j + "']", ExcelUtil.GetCellData(row, buildcol, sheetname), Property_type.XPath);
                    Console.WriteLine("Decription details entered");


                    //Perform.SelectDropDown(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_LocationList_Repeater1_ctl_BOP_Location_"+i+"_ctl_BOP_BuildingList_"+i+"_Repeater1_"+i+"_ctl_BOP_Building_"+j+"_ctl_BOP_Building_Information_"+j+"_ddlOccupancy_"+j+"']", ExcelUtil.GetCellData(row, buildcol + 1, sheetname), Property_type.XPath);
                    Perform.selectbyvalue(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_LocationList_Repeater1_ctl_BOP_Location_" + i + "_ctl_BOP_BuildingList_" + i + "_Repeater1_" + i + "_ctl_BOP_Building_" + j + "_ctl_BOP_Building_Information_" + j + "_ddlOccupancy_" + j + "']", ExcelUtil.GetCellData(row, buildcol + 1, sheetname));
                    Console.WriteLine("Occupancy details entered");
                    Perform.Wait();
                    Perform.SelectDropDown(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_LocationList_Repeater1_ctl_BOP_Location_" + i + "_ctl_BOP_BuildingList_" + i + "_Repeater1_" + i + "_ctl_BOP_Building_" + j + "_ctl_BOP_Building_Information_" + j + "_ddlConstruction_" + j + "']", ExcelUtil.GetCellData(row, buildcol + 2, sheetname), Property_type.XPath);
                    Console.WriteLine("Construction details entered");
                    Perform.SelectDropDown(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_LocationList_Repeater1_ctl_BOP_Location_" + i + "_ctl_BOP_BuildingList_" + i + "_Repeater1_" + i + "_ctl_BOP_Building_" + j + "_ctl_BOP_Building_Information_" + j + "_ddlAutomaticIncrease_" + j + "']", ExcelUtil.GetCellData(row, buildcol + 3, sheetname), Property_type.XPath);
                    // Perform.selectbyvalue(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_LocationList_Repeater1_ctl_BOP_Location_" + i + "_ctl_BOP_BuildingList_" + i + "_Repeater1_" + i + "_ctl_BOP_Building_" + j + "_ctl_BOP_Building_Information_" + j + "_ddlAutomaticIncrease_" + j + "']", ExcelUtil.GetCellData(row, buildcol + 3, sheetname));
                    if (Perform.driver.FindElement(By.XPath(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_LocationList_Repeater1_ctl_BOP_Location_" + i + "_ctl_BOP_BuildingList_" + i + "_Repeater1_" + i + "_ctl_BOP_Building_" + j + "_ctl_BOP_Building_Information_" + j + "_ddlPropertyDeductible_" + j + "']")).Enabled)
                    {
                        Perform.SelectDropDown(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_LocationList_Repeater1_ctl_BOP_Location_" + i + "_ctl_BOP_BuildingList_" + i + "_Repeater1_" + i + "_ctl_BOP_Building_" + j + "_ctl_BOP_Building_Information_" + j + "_ddlPropertyDeductible_" + j + "']", ExcelUtil.GetCellData(row, buildcol + 4, sheetname), Property_type.XPath);
                    }
                    Perform.EnterText(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_LocationList_Repeater1_ctl_BOP_Location_" + i + "_ctl_BOP_BuildingList_" + i + "_Repeater1_" + i + "_ctl_BOP_Building_" + j + "_ctl_BOP_Building_Information_" + j + "_txtBuildingLimit_" + j + "']", ExcelUtil.GetCellData(row, buildcol + 5, sheetname), Property_type.XPath);
                    Console.WriteLine("Building limit entered");


                    Perform.SelectDropDown(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_LocationList_Repeater1_ctl_BOP_Location_" + i + "_ctl_BOP_BuildingList_" + i + "_Repeater1_" + i + "_ctl_BOP_Building_" + j + "_ctl_BOP_Building_Information_" + j + "_ddlBuildingValuation_" + j + "']", ExcelUtil.GetCellData(row, buildcol + 6, sheetname), Property_type.XPath);
                    Console.WriteLine("Building valuation entered");

                    if (ExcelUtil.GetCellData(row, buildcol + 7, sheetname) == "YES")
                    {
                        Perform.Click(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_LocationList_Repeater1_ctl_BOP_Location_" + i + "_ctl_BOP_BuildingList_" + i + "_Repeater1_" + i + "_ctl_BOP_Building_" + j + "_ctl_BOP_Building_Information_" + j + "_chkACVRoofing_" + j + "']", Property_type.XPath);
                    }

                    if (ExcelUtil.GetCellData(row, buildcol + 8, sheetname) == "YES")
                        Perform.Click(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_LocationList_Repeater1_ctl_BOP_Location_" + i + "_ctl_BOP_BuildingList_" + i + "_Repeater1_" + i + "_ctl_BOP_Building_" + j + "_ctl_BOP_Building_Information_" + j + "_chkBuildingValuationIncludedInBlanketRating_" + j + "']", Property_type.XPath);
                    if (ExcelUtil.GetCellData(row, buildcol + 9, sheetname) == "YES")
                        Perform.Click(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_LocationList_Repeater1_ctl_BOP_Location_" + i + "_ctl_BOP_BuildingList_" + i + "_Repeater1_" + i + "_ctl_BOP_Building_" + j + "_ctl_BOP_Building_Information_" + j + "_chkMineSubsidence_" + j + "']", Property_type.XPath);
                    if (ExcelUtil.GetCellData(row, buildcol + 10, sheetname) == "YES")
                        Perform.Click(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_LocationList_Repeater1_ctl_BOP_Location_" + i + "_ctl_BOP_BuildingList_" + i + "_Repeater1_" + i + "_ctl_BOP_Building_" + j + "_ctl_BOP_Building_Information_" + j + "_chkSprinklered_" + j + "']", Property_type.XPath);
                    if (ExcelUtil.GetCellData(row, buildcol + 11, sheetname) != "")
                    {
                        Perform.EnterText(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_LocationList_Repeater1_ctl_BOP_Location_" + i + "_ctl_BOP_BuildingList_" + i + "_Repeater1_" + i + "_ctl_BOP_Building_" + j + "_ctl_BOP_Building_Information_" + j + "_txtPersonalPropertyLimit_" + j + "']", ExcelUtil.GetCellData(row, buildcol + 11, sheetname), Property_type.XPath);
                    }
                    else
                    {
                        Console.WriteLine("No Personal property value");
                    }
                    Perform.SelectDropDown(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_LocationList_Repeater1_ctl_BOP_Location_" + i + "_ctl_BOP_BuildingList_" + i + "_Repeater1_" + i + "_ctl_BOP_Building_" + j + "_ctl_BOP_Building_Information_" + j + "_ddlValuationMethod_" + j + "']", ExcelUtil.GetCellData(row, buildcol + 12, sheetname), Property_type.XPath);
                    if (ExcelUtil.GetCellData(row, buildcol + 13, sheetname) == "YES")
                        Perform.Click(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_LocationList_Repeater1_ctl_BOP_Location_" + i + "_ctl_BOP_BuildingList_" + i + "_Repeater1_" + i + "_ctl_BOP_Building_" + j + "_ctl_BOP_Building_Information_" + j + "_chkValuationMethodIncludedInBlanketRating_" + j + "']", Property_type.XPath);
                    if (ExcelUtil.GetCellData(row, buildcol + 14, sheetname) != "")
                    {
                        Perform.EnterText(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_LocationList_Repeater1_ctl_BOP_Location_" + i + "_ctl_BOP_BuildingList_" + i + "_Repeater1_" + i + "_ctl_BOP_Building_" + j + "_ctl_BOP_Building_Information_" + j + "_txtFeetToHydrant_" + j + "']", ExcelUtil.GetCellData(row, buildcol + 14, sheetname), Property_type.XPath);
                    }
                    else
                    {
                        Console.WriteLine("Feet to Hydrant value is not present");
                    }
                    if (ExcelUtil.GetCellData(row, buildcol + 15, sheetname) != "")
                    {
                        Perform.EnterText(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_LocationList_Repeater1_ctl_BOP_Location_" + i + "_ctl_BOP_BuildingList_" + i + "_Repeater1_" + i + "_ctl_BOP_Building_" + j + "_ctl_BOP_Building_Information_" + j + "_txtMilesToFireDepartment_" + j + "']", ExcelUtil.GetCellData(row, buildcol + 15, sheetname), Property_type.XPath);
                    }
                    else
                    {
                        Console.WriteLine("Miles to Fire Station value is not present");
                    }
                    Perform.SelectDropDown(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_LocationList_Repeater1_ctl_BOP_Location_" + i + "_ctl_BOP_BuildingList_" + i + "_Repeater1_" + i + "_ctl_BOP_Building_" + j + "_ctl_BOP_Building_Information_" + j + "_ddlProtectionClass_" + j + "']", ExcelUtil.GetCellData(row, buildcol + 16, sheetname), Property_type.XPath);
                    Console.WriteLine("Protection class is selected");
                    Perform.SelectDropDown(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_LocationList_Repeater1_ctl_BOP_Location_" + i + "_ctl_BOP_BuildingList_" + i + "_Repeater1_" + i + "_ctl_BOP_Building_" + j + "_ctl_BOP_Building_Information_" + j + "_ctlClassificationsList_" + j + "_Repeater1_" + j + "_ctl_BuildingClassificationItem_0_ddlProgram_0']", ExcelUtil.GetCellData(row, buildcol + 18, sheetname), Property_type.XPath);
                    Console.WriteLine("Building Program selected");
                    System.Threading.Thread.Sleep(1000);
                    Perform.waitTillElementToAppear(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_LocationList_Repeater1_ctl_BOP_Location_" + i + "_ctl_BOP_BuildingList_" + i + "_Repeater1_" + i + "_ctl_BOP_Building_" + j + "_ctl_BOP_Building_Information_" + j + "_ctlClassificationsList_" + j + "_Repeater1_" + j + "_ctl_BuildingClassificationItem_0_ddlClassification_0']");
                    Perform.selectbyvalue("//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_LocationList_Repeater1_ctl_BOP_Location_" + i + "_ctl_BOP_BuildingList_" + i + "_Repeater1_" + i + "_ctl_BOP_Building_" + j + "_ctl_BOP_Building_Information_" + j + "_ctlClassificationsList_" + j + "_Repeater1_" + j + "_ctl_BuildingClassificationItem_0_ddlClassification_0']", ExcelUtil.GetCellData(row, buildcol + 19, sheetname));
                    //Perform.SelectDropDown(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_LocationList_Repeater1_ctl_BOP_Location_" + i + "_ctl_BOP_BuildingList_" + i + "_Repeater1_" + i + "_ctl_BOP_Building_" + j + "_ctl_BOP_Building_Information_" + j + "_ctlClassificationsList_" + j + "_Repeater1_" + j + "_ctl_BuildingClassificationItem_" + j + "_ddlClassification_" + j + "']", ExcelUtil.GetCellData(row, buildcol + 19, sheetname), Property_type.XPath);
                    Console.WriteLine("Classification entered");
                    if (ExcelUtil.GetCellData(row, buildcol + 18, sheetname) == "Motel" || ExcelUtil.GetCellData(row, buildcol + 18, sheetname) == "Office" || ExcelUtil.GetCellData(row, buildcol + 18, sheetname) == "Contractors - Shop")
                    {
                        Perform.waitTillElementToAppear("html / body / div[1] / div[3] / div / button");
                        // Property_Collection.driver.SwitchTo().Frame("html/body/div[1]/div[3]/div/button");
                        Perform.Click("html / body / div[1] / div[3] / div / button", Property_type.XPath);
                    }

                    //Perform.EnterText(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_LocationList_Repeater1_ctl_BOP_Location_" + i + "_ctl_BOP_BuildingList_" + i + "_Repeater1_" + i + "_ctl_BOP_Building_" + j + "_ctl_BOP_Building_Information_" + j + "_ctlClassificationsList_" + j + "_Repeater1_" + j + "_ctl_BuildingClassificationItem_" + j + "_txtClassCode_" + j + "']", ExcelUtil.GetCellData(row, buildcol + 20, sheetname), Property_type.XPath);
                    if (ExcelUtil.GetCellData(row, buildcol + 21, sheetname) != "")
                    { 
                    if (Perform.driver.FindElement(By.XPath(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_LocationList_Repeater1_ctl_BOP_Location_" + i + "_ctl_BOP_BuildingList_" + i + "_Repeater1_" + i + "_ctl_BOP_Building_" + j + "_ctl_BOP_Building_Information_" + j + "_ctlClassificationsList_" + j + "_Repeater1_" + j + "_ctl_BuildingClassificationItem_" + j + "_txtAnnualReceipts_" + j + "']")).Displayed)
                    {
                        Perform.EnterText(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_LocationList_Repeater1_ctl_BOP_Location_" + i + "_ctl_BOP_BuildingList_" + i + "_Repeater1_" + i + "_ctl_BOP_Building_" + j + "_ctl_BOP_Building_Information_" + j + "_ctlClassificationsList_" + j + "_Repeater1_" + j + "_ctl_BuildingClassificationItem_" + j + "_txtAnnualReceipts_" + j + "']", ExcelUtil.GetCellData(row, buildcol + 21, sheetname), Property_type.XPath);
                    }
                    else
                        {
                            Console.WriteLine("No annual receipts");
                        }
                    
                }

                    if (ExcelUtil.GetCellData(row, buildcol + 22, sheetname) != "")
                    {
                        if (Perform.driver.FindElement(By.XPath(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_LocationList_Repeater1_ctl_BOP_Location_" + i + "_ctl_BOP_BuildingList_" + i + "_Repeater1_" + i + "_ctl_BOP_Building_" + j + "_ctl_BOP_Building_Information_" + j + "_ctlClassificationsList_" + j + "_Repeater1_" + j + "_ctl_BuildingClassificationItem_" + j + "_txtEmployeePayroll_" + j + "']")).Displayed)
                        {
                            Perform.EnterText(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_LocationList_Repeater1_ctl_BOP_Location_" + i + "_ctl_BOP_BuildingList_" + i + "_Repeater1_" + i + "_ctl_BOP_Building_" + j + "_ctl_BOP_Building_Information_" + j + "_ctlClassificationsList_" + j + "_Repeater1_" + j + "_ctl_BuildingClassificationItem_" + j + "_txtEmployeePayroll_" + j + "']", ExcelUtil.GetCellData(row, buildcol + 22, sheetname), Property_type.XPath);
                        }
                        else
                        {
                            Console.WriteLine("No employee payroll");
                        }


                        if (ExcelUtil.GetCellData(row, buildcol + 23, sheetname) != "")
                        {
                            if (Perform.driver.FindElement(By.XPath(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_LocationList_Repeater1_ctl_BOP_Location_" + i + "_ctl_BOP_BuildingList_" + i + "_Repeater1_" + i + "_ctl_BOP_Building_" + j + "_ctl_BOP_Building_Information_" + j + "_ctlClassificationsList_" + j + "_Repeater1_" + j + "_ctl_BuildingClassificationItem_" + j + "_txtNumOfficersPartnersIndInsureds_" + j + "']")).Displayed)
                            {
                                Perform.EnterText(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_LocationList_Repeater1_ctl_BOP_Location_" + i + "_ctl_BOP_BuildingList_" + i + "_Repeater1_" + i + "_ctl_BOP_Building_" + j + "_ctl_BOP_Building_Information_" + j + "_ctlClassificationsList_" + j + "_Repeater1_" + j + "_ctl_BuildingClassificationItem_" + j + "_txtNumOfficersPartnersIndInsureds_" + j + "']", ExcelUtil.GetCellData(row, buildcol + 23, sheetname), Property_type.XPath);
                            }
                            else
                            {
                                Console.WriteLine("No insured partners");
                            }
                        }
                    }
                    Console.WriteLine(row+ buildcol+ sheetname);
                    if (ExcelUtil.GetCellData(row, buildcol + 24, sheetname) == "YES")
                        {
                        
                            Perform.Click(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_LocationList_Repeater1_ctl_BOP_Location_" + i + "_ctl_BOP_BuildingList_" + i + "_Repeater1_" + i + "_ctl_BOP_Building_" + j + "_ctl_BOP_Building_Information_" + j + "_ctlClassificationsList_" + j + "_Repeater1_" + j + "_ctl_BuildingClassificationItem_0_chkPrimaryClassification_0']", Property_type.XPath);
                        Perform.Wait();
                        Console.WriteLine("Primary Classification selected");
                        }
                    Perform.Wait();
                        //Check for optional building coverage details
                        //Accounts Receivable - On Premises
                        if (ExcelUtil.GetCellData(row, buildcol + 25, sheetname) == "YES")
                        {
                            Perform.Click(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_LocationList_Repeater1_ctl_BOP_Location_" + i + "_ctl_BOP_BuildingList_" + i + "_Repeater1_" + i + "_ctl_BOP_Building_" + j + "_ctl_BOP_BuildingCoverages_" + j + "_chkAcctsReceivableONPremises_" + j + "']", Property_type.XPath);
                            Console.WriteLine("Accounts receivables clicked");
                            Perform.EnterText(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_LocationList_Repeater1_ctl_BOP_Location_" + i + "_ctl_BOP_BuildingList_" + i + "_Repeater1_" + i + "_ctl_BOP_Building_" + j + "_ctl_BOP_BuildingCoverages_" + j + "_txtAcctsReceivableOnPremisesTotalLimit_" + j + "']", ExcelUtil.GetCellData(row, buildcol + 26, sheetname), Property_type.XPath);
                            Console.WriteLine("Accounts receivables entered");
                        }
                        else
                        {
                            Console.WriteLine("Accounts receivable is not selected");
                        }
                        //Valuable Papers - On Premises
                        if (ExcelUtil.GetCellData(row, buildcol + 27, sheetname) == "YES")
                        {
                            Perform.Click(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_LocationList_Repeater1_ctl_BOP_Location_" + i + "_ctl_BOP_BuildingList_" + i + "_Repeater1_" + i + "_ctl_BOP_Building_" + j + "_ctl_BOP_BuildingCoverages_" + j + "_chkValuablePapersOnPremises_" + j + "']", Property_type.XPath);
                            Perform.EnterText(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_LocationList_Repeater1_ctl_BOP_Location_" + i + "_ctl_BOP_BuildingList_" + i + "_Repeater1_" + i + "_ctl_BOP_Building_" + j + "_ctl_BOP_BuildingCoverages_" + j + "_txtValuablePapersOnPremisesTotalLimit_" + j + "']", ExcelUtil.GetCellData(row, buildcol + 28, sheetname), Property_type.XPath);
                            Console.WriteLine("Valuable Papers - On Premises entered");
                        }
                        else
                        {
                            Console.WriteLine("Valuable Papers is not selected");
                        }

                        //Ordinance or Law
                        if (ExcelUtil.GetCellData(row, buildcol + 29, sheetname) == "YES")
                        {
                            Perform.Click(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_LocationList_Repeater1_ctl_BOP_Location_" + i + "_ctl_BOP_BuildingList_" + i + "_Repeater1_" + i + "_ctl_BOP_Building_" + j + "_ctl_BOP_BuildingCoverages_" + j + "_chkOrdinanceOrLaw_" + j + "']", Property_type.XPath);
                            Perform.EnterText(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_LocationList_Repeater1_ctl_BOP_Location_" + i + "_ctl_BOP_BuildingList_" + i + "_Repeater1_" + i + "_ctl_BOP_Building_" + j + "_ctl_BOP_BuildingCoverages_" + j + "_txtDemolitionCostLimit_" + j + "']", ExcelUtil.GetCellData(row, buildcol + 30, sheetname), Property_type.XPath);
                            Perform.EnterText(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_LocationList_Repeater1_ctl_BOP_Location_" + i + "_ctl_BOP_BuildingList_" + i + "_Repeater1_" + i + "_ctl_BOP_Building_" + j + "_ctl_BOP_BuildingCoverages_" + j + "_txtIncreasedCostOfConstructionLimit_" + j + "']", ExcelUtil.GetCellData(row, buildcol + 31, sheetname), Property_type.XPath);
                            Perform.EnterText(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_LocationList_Repeater1_ctl_BOP_Location_" + i + "_ctl_BOP_BuildingList_" + i + "_Repeater1_" + i + "_ctl_BOP_Building_" + j + "_ctl_BOP_BuildingCoverages_" + j + "_txtDemolitionAndIncreasedCostCombinedLimit_" + j + "']", ExcelUtil.GetCellData(row, buildcol + 32, sheetname), Property_type.XPath);
                        }
                        else
                        {
                            Console.WriteLine("Ordinance or Law is not selected");
                        }
                        //Spoilage
                        if (ExcelUtil.GetCellData(row, buildcol + 33, sheetname) == "YES")
                        {
                            Perform.Click(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_LocationList_Repeater1_ctl_BOP_Location_" + i + "_ctl_BOP_BuildingList_" + i + "_Repeater1_" + i + "_ctl_BOP_Building_" + j + "_ctl_BOP_BuildingCoverages_" + j + "_chkSpoilage_" + j + "']", Property_type.XPath);
                            Perform.selectbyvalue(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_LocationList_Repeater1_ctl_BOP_Location_" + i + "_ctl_BOP_BuildingList_" + i + "_Repeater1_" + i + "_ctl_BOP_Building_" + j + "_ctl_BOP_BuildingCoverages_" + j + "_ddlSpoilagePropertyClassification_" + j + "']", ExcelUtil.GetCellData(row, buildcol + 34, sheetname));
                            Perform.EnterText(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_LocationList_Repeater1_ctl_BOP_Location_" + i + "_ctl_BOP_BuildingList_" + i + "_Repeater1_" + i + "_ctl_BOP_Building_" + j + "_ctl_BOP_BuildingCoverages_" + j + "_txtSpoilageTotalLimit_" + j + "']", ExcelUtil.GetCellData(row, buildcol + 35, sheetname), Property_type.XPath);
                            if (ExcelUtil.GetCellData(row, buildcol + 36, sheetname) == "YES")
                                Perform.Click(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_LocationList_Repeater1_ctl_BOP_Location_" + i + "_ctl_BOP_BuildingList_" + i + "_Repeater1_" + i + "_ctl_BOP_Building_" + j + "_ctl_BOP_BuildingCoverages_" + j + "_chkRefrigeratorMaintenanceAgreement_" + j + "']", Property_type.XPath);
                            if (ExcelUtil.GetCellData(row, buildcol + 37, sheetname) == "YES")
                                Perform.Click(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_LocationList_Repeater1_ctl_BOP_Location_"+i+"_ctl_BOP_BuildingList_"+i+"_Repeater1_"+i+"_ctl_BOP_Building_"+j+"_ctl_BOP_BuildingCoverages_"+j+"_chkBreakdownOrContamination_"+j+"']", Property_type.XPath);
                            if (ExcelUtil.GetCellData(row, buildcol + 38, sheetname) == "YES")
                                Perform.Click(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_LocationList_Repeater1_ctl_BOP_Location_" + i + "_ctl_BOP_BuildingList_" + i + "_Repeater1_" + i + "_ctl_BOP_Building_" + j + "_ctl_BOP_BuildingCoverages_" + j + "_chkPowerOutage_" + j + "']", Property_type.XPath);
                        }
                        else
                        {
                            Console.WriteLine("Spoilage is not selected");
                        }
                    if (ExcelUtil.GetCellData(row, buildcol + 47, sheetname) == "YES")
                    {
                        if (Perform.driver.FindElement(By.XPath(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_LocationList_Repeater1_ctl_BOP_Location_" + i + "_ctl_BOP_BuildingList_" + i + "_Repeater1_" + i + "_ctl_BOP_Building_" + j + "_ctl_BOP_BuildingCoverages_" + j + "_chkSelfStorage_" + j + "']")).Enabled)
                        {
                            Perform.Click(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_LocationList_Repeater1_ctl_BOP_Location_" + i + "_ctl_BOP_BuildingList_" + i + "_Repeater1_" + i + "_ctl_BOP_Building_" + j + "_ctl_BOP_BuildingCoverages_" + j + "_chkSelfStorage_" + j + "']", Property_type.XPath);
                            Perform.EnterText(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_LocationList_Repeater1_ctl_BOP_Location_" + i + "_ctl_BOP_BuildingList_" + i + "_Repeater1_" + i + "_ctl_BOP_Building_" + j + "_ctl_BOP_BuildingCoverages_" + j + "_txtStorageLimit_" + j + "']", ExcelUtil.GetCellData(row, buildcol + 48, sheetname), Property_type.XPath);
                        }
                    }
                        //Motel
                        if (ExcelUtil.GetCellData(row, buildcol + 49, sheetname) == "YES")
                        {
                            Perform.Click(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_LocationList_Repeater1_ctl_BOP_Location_" + i + "_ctl_BOP_BuildingList_" + i + "_Repeater1_" + i + "_ctl_BOP_Building_" + j + "_ctl_BOP_BuildingCoverages_" + j + "_chkMotel_" + j + "']", Property_type.XPath);
                            Perform.SelectDropDown(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_LocationList_Repeater1_ctl_BOP_Location_" + i + "_ctl_BOP_BuildingList_" + i + "_Repeater1_" + i + "_ctl_BOP_Building_" + j + "_ctl_BOP_BuildingCoverages_" + j + "_ddlMotelLiabilityLimit_" + j + "']", ExcelUtil.GetCellData(row, buildcol + 50, sheetname), Property_type.XPath);
                            Perform.SelectDropDown(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_LocationList_Repeater1_ctl_BOP_Location_" + i + "_ctl_BOP_BuildingList_" + i + "_Repeater1_" + i + "_ctl_BOP_Building_" + j + "_ctl_BOP_BuildingCoverages_" + j + "_ddlMotelSafeDepositBoxLimit_" + j + "']", ExcelUtil.GetCellData(row, buildcol + 51, sheetname), Property_type.XPath);
                            Perform.SelectDropDown(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_LocationList_Repeater1_ctl_BOP_Location_" + i + "_ctl_BOP_BuildingList_" + i + "_Repeater1_" + i + "_ctl_BOP_Building_" + j + "_ctl_BOP_BuildingCoverages_" + j + "_ddlMotelSafeDepositBoxDeductible_" + j + "']", ExcelUtil.GetCellData(row, buildcol + 52, sheetname), Property_type.XPath);
                        }
                        else
                        {
                            Console.WriteLine("Motel is not selected");
                        }
                    Perform.Wait();
                   
                        //Funeral
                    if (ExcelUtil.GetCellData(row, buildcol + 53, sheetname) == "YES")
                    {
                        if (Perform.driver.FindElement(By.XPath(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_LocationList_Repeater1_ctl_BOP_Location_" + i + "_ctl_BOP_BuildingList_" + i + "_Repeater1_" + i + "_ctl_BOP_Building_" + j + "_ctl_BOP_BuildingCoverages_" + j + "_chkFuneralDirectors_" + j + "']")).Enabled)
                        {
                            Perform.Click(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_LocationList_Repeater1_ctl_BOP_Location_" + i + "_ctl_BOP_BuildingList_" + i + "_Repeater1_" + i + "_ctl_BOP_Building_" + j + "_ctl_BOP_BuildingCoverages_" + j + "_chkFuneralDirectors_" + j + "']", Property_type.XPath);
                            Perform.EnterText(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_LocationList_Repeater1_ctl_BOP_Location_" + i + "_ctl_BOP_BuildingList_" + i + "_Repeater1_" + i + "_ctl_BOP_Building_" + j + "_ctl_BOP_BuildingCoverages_0_txtFuneralNumOfEmployees_0']", ExcelUtil.GetCellData(row, buildcol + 54, sheetname), Property_type.XPath);
                        }
                        else
                        {
                            Console.WriteLine("Funeral checkbox is disabled");
                        }
                        
                    }
                        Console.WriteLine("Location " + (i+ 1) + " && Building" + (j+1) + " details entered");
                 if (i > 0)
                    {
                      Perform.driver.FindElement(By.XPath(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_LocationList_Repeater1_ctl_BOP_Location_" + i + "_ctl_BOP_BuildingList_" + i + "_Repeater1_" + i + "_ctl_BOP_Building_" + j + "_btnSave_" + j + "']")).SendKeys(Keys.PageDown);
                        Perform.waitTillElementToAppear(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_LocationList_Repeater1_ctl_BOP_Location_" + i + "_ctl_BOP_BuildingList_" + i + "_Repeater1_" + i + "_ctl_BOP_Building_" + j + "_btnSave_" + j + "']");
                   }
                    Perform.Click(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_LocationList_Repeater1_ctl_BOP_Location_"+i+ "_ctl_BOP_BuildingList_" + i + "_Repeater1_" + i + "_ctl_BOP_Building_" + j + "_btnSave_" + j + "']", Property_type.XPath);
                  //  Console.WriteLine("Save building is clicked");
                    //System.Threading.Thread.Sleep(1000);
                        buildcol = buildcol + 55;
                    }


                    

                    
                
            }
        }
        public static void IRPMValue(int row, string sheetname)
        {
            Perform.waitTillElementToAppear(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_IRPM_btnSubmitRate']");
            int col = 1;
            for (int i = 0; i < 8; i++)
            {
                if (ExcelUtil.GetCellData(row, col, sheetname) != "")
                {
                    Perform.Click(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_IRPM_rptIRPM_txtRisk_" + i + "']", Property_type.XPath);
                    Console.WriteLine(ExcelUtil.GetCellData(row, col, sheetname).Substring(0, 1));
                    if (ExcelUtil.GetCellData(row, col, sheetname).Substring(0,1)== "+")
                    {
                        string strinc = ExcelUtil.GetCellData(row, col, sheetname);
                        int inc = Int32.Parse(strinc.Substring(1,strinc.Length-1 ));
                        Console.WriteLine(inc.ToString());
                        for (int j=0;j<inc;j++)
                        {
                            Perform.driver.FindElement(By.XPath(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_IRPM_rptIRPM_txtRisk_" + i + "']")).SendKeys(Keys.ArrowUp);  
                        }

                    }
                    if (ExcelUtil.GetCellData(row, col, sheetname).Substring(0, 1) == "-")
                    {
                        string strinc = ExcelUtil.GetCellData(row, col, sheetname);
                        int inc = Int32.Parse(strinc.Substring(1, strinc.Length - 1));
                        Console.WriteLine(inc.ToString());
                        for (int j = 0; j < inc; j++)
                        {
                            Perform.driver.FindElement(By.XPath(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_IRPM_rptIRPM_txtRisk_" + i + "']")).SendKeys(Keys.ArrowDown);
                        }

                    }
                    Perform.Click(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_IRPM_tblMain']/tbody/tr/td", Property_type.XPath);
                    Perform.waitTillElementToAppear(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_IRPM_rptIRPM_txtIRPMDescription_" + i + "']");
                    Perform.EnterText(".//*[@id='cphMain_ctl_WorkflowManager_BOP_Quote_ctl_BOP_IRPM_rptIRPM_txtIRPMDescription_" + i + "']", ExcelUtil.GetCellData(row, col + 1, sheetname), Property_type.XPath);
                }
                col = col + 2;
            }
            
        }

        public static void Additonal_App(int row, string addapp, string sheetname)
        {
            
           
            for (int i = 0; i < Int32.Parse(addapp); i++)
            {
                sheetname = "Application_Location_" + i;
                
             
                if (ExcelUtil.GetCellData(row, 3, sheetname)=="YES")
                {
                    Perform.Click(".//*[@id='cphMain_ctl_WorkflowManager_App_BOP_ctl_AppSection_BOP_ctl_AddlPolicyholderList_btnAddAdditionalPolicyholder']", Property_type.XPath);
                    Perform.SelectDropDown(".//*[@id='cphMain_ctl_WorkflowManager_App_BOP_ctl_AppSection_BOP_ctl_AddlPolicyholderList_Repeater1_ctl_App_APH_" + i + "_ddTaxIdType_" + i + "']", ExcelUtil.GetCellData(row, 4, sheetname), Property_type.XPath);
                    Perform.EnterText(".//*[@id='cphMain_ctl_WorkflowManager_App_BOP_ctl_AppSection_BOP_ctl_AddlPolicyholderList_Repeater1_ctl_App_APH_" + i + "_txtName_" + i + "']", ExcelUtil.GetCellData(row, 5, sheetname), Property_type.XPath);
                    if (ExcelUtil.GetCellData(row, 4, sheetname) == "FEIN")
                    {
                        Perform.waitTillElementToAppear(".//*[@id='cphMain_ctl_WorkflowManager_App_BOP_ctl_AppSection_BOP_ctl_AddlPolicyholderList_Repeater1_ctl_App_APH_" + i + "_txtFEIN_" + i + "']");
                        Perform.EnterText(".//*[@id='cphMain_ctl_WorkflowManager_App_BOP_ctl_AppSection_BOP_ctl_AddlPolicyholderList_Repeater1_ctl_App_APH_" + i + "_txtFEIN_" + i + "']", ExcelUtil.GetCellData(row, 6, sheetname), Property_type.XPath);
                    }
                    if (ExcelUtil.GetCellData(row, 4, sheetname) == "SSN")
                    {
                        Perform.waitTillElementToAppear(".//*[@id='cphMain_ctl_WorkflowManager_App_BOP_ctl_AppSection_BOP_ctl_AddlPolicyholderList_Repeater1_ctl_App_APH_" + i + "_txtSSN_" + i + "']");
                        Perform.EnterText(".//*[@id='cphMain_ctl_WorkflowManager_App_BOP_ctl_AppSection_BOP_ctl_AddlPolicyholderList_Repeater1_ctl_App_APH_" + i + "_txtSSN_" + i + "']", ExcelUtil.GetCellData(row, 6, sheetname), Property_type.XPath);
                    }
                }
                else
                    Console.WriteLine("No additional applicants");
            }
            string nooflocation = ExcelUtil.GetCellData(row, 3, "TestCase");
        
            for (int i = 0; i < Int32.Parse(nooflocation); i++)
            {
                sheetname = "Application_Location_" + i;
                if (i > 0)
                {
                    Perform.Click(".//*[@id='cphMain_ctl_WorkflowManager_App_BOP_ctl_AppSection_BOP_ctl_LocationList_Repeater1_ctl_BOP_App_Location_" + i + "_lblAccordHeader_" + i + "']", Property_type.XPath);
                }
                string nofbuilding = ExcelUtil.GetCellData(row, 4+i, "TestCase");
                int col = 19;
                for (int j = 0; j < Int32.Parse(nofbuilding); j++)
                {
                   
                    if (j>0)
                    {
                        //Perform.Click(".//*[@id='ui-id-"+spval+"']/span[1]", Property_type.XPath);
                  
                        Perform.Click(".//*[@id='cphMain_ctl_WorkflowManager_App_BOP_ctl_AppSection_BOP_ctl_LocationList_Repeater1_ctl_BOP_App_Location_"+i+"_ctl_BOP_App_BuildingList_"+i+"_Repeater1_"+i+"_ctl_BOP_App_Building_"+j+"_lblAccordHeader_"+j+"']", Property_type.XPath);
                        Console.WriteLine("Span button clicked");
                    }
                    Perform.waitTillElementToAppear(".//*[@id='cphMain_ctl_WorkflowManager_App_BOP_ctl_AppSection_BOP_ctl_LocationList_Repeater1_ctl_BOP_App_Location_" + i + "_ctl_BOP_App_BuildingList_" + i + "_Repeater1_" + i + "_ctl_BOP_App_Building_" + j + "_txtSquareFeet_" + j + "']");
                    Console.WriteLine(sheetname);
                    Console.WriteLine(ExcelUtil.GetCellData(row, col, sheetname));
                    Perform.EnterText(".//*[@id='cphMain_ctl_WorkflowManager_App_BOP_ctl_AppSection_BOP_ctl_LocationList_Repeater1_ctl_BOP_App_Location_" + i + "_ctl_BOP_App_BuildingList_" + i + "_Repeater1_" + i + "_ctl_BOP_App_Building_" + j + "_txtSquareFeet_" + j + "']", ExcelUtil.GetCellData(row, col, sheetname), Property_type.XPath);
                    Perform.EnterText(".//*[@id='cphMain_ctl_WorkflowManager_App_BOP_ctl_AppSection_BOP_ctl_LocationList_Repeater1_ctl_BOP_App_Location_" + i + "_ctl_BOP_App_BuildingList_" + i + "_Repeater1_" + i + "_ctl_BOP_App_Building_" + j + "_txtYearRoofUpdated_" + j + "']", ExcelUtil.GetCellData(row, col + 1, sheetname), Property_type.XPath);
                    Perform.EnterText(".//*[@id='cphMain_ctl_WorkflowManager_App_BOP_ctl_AppSection_BOP_ctl_LocationList_Repeater1_ctl_BOP_App_Location_" + i + "_ctl_BOP_App_BuildingList_" + i + "_Repeater1_" + i + "_ctl_BOP_App_Building_" + j + "_txtYearPlumbingUpdated_" + j + "']", ExcelUtil.GetCellData(row, col + 2, sheetname), Property_type.XPath);
                    Perform.EnterText(".//*[@id='cphMain_ctl_WorkflowManager_App_BOP_ctl_AppSection_BOP_ctl_LocationList_Repeater1_ctl_BOP_App_Location_" + i + "_ctl_BOP_App_BuildingList_" + i + "_Repeater1_" + i + "_ctl_BOP_App_Building_" + j + "_txtYearBuilt_" + j + "']", ExcelUtil.GetCellData(row, col + 3, sheetname), Property_type.XPath);
                    Perform.EnterText(".//*[@id='cphMain_ctl_WorkflowManager_App_BOP_ctl_AppSection_BOP_ctl_LocationList_Repeater1_ctl_BOP_App_Location_" + i + "_ctl_BOP_App_BuildingList_" + i + "_Repeater1_" + i + "_ctl_BOP_App_Building_" + j + "_txtYearWiringUpdated_" + j + "']", ExcelUtil.GetCellData(row, col + 4, sheetname), Property_type.XPath);
                    Perform.EnterText(".//*[@id='cphMain_ctl_WorkflowManager_App_BOP_ctl_AppSection_BOP_ctl_LocationList_Repeater1_ctl_BOP_App_Location_" + i + "_ctl_BOP_App_BuildingList_" + i + "_Repeater1_" + i + "_ctl_BOP_App_Building_" + j + "_txtYearHeatUpdated_" + j + "']", ExcelUtil.GetCellData(row, col + 5, sheetname), Property_type.XPath);

                    if (ExcelUtil.GetCellData(row, col + 6, sheetname) != "")
                    {
                        Perform.SelectDropDown(".//*[@id='cphMain_ctl_WorkflowManager_App_BOP_ctl_AppSection_BOP_ctl_LocationList_Repeater1_ctl_BOP_App_Location_" + i + "_ctl_BOP_App_BuildingList_" + i + "_Repeater1_" + i + "_ctl_BOP_App_Building_" + j + "_ddlBuildingLimitLossPayeeName_" + j + "']", ExcelUtil.GetCellData(row, col + 6, sheetname), Property_type.XPath);
                        Perform.SelectDropDown(".//*[@id='cphMain_ctl_WorkflowManager_App_BOP_ctl_AppSection_BOP_ctl_LocationList_Repeater1_ctl_BOP_App_Location_" + i + "_ctl_BOP_App_BuildingList_" + i + "_Repeater1_" + i + "_ctl_BOP_App_Building_" + j + "_ddlBuildingLimitLossPayeeType_" + j + "']", ExcelUtil.GetCellData(row, col + 7, sheetname), Property_type.XPath);
                        Perform.SelectDropDown(".//*[@id='cphMain_ctl_WorkflowManager_App_BOP_ctl_AppSection_BOP_ctl_LocationList_Repeater1_ctl_BOP_App_Location_" + i + "_ctl_BOP_App_BuildingList_" + i + "_Repeater1_" + i + "_ctl_BOP_App_Building_" + j + "_ddlBuildingLimitATMA_" + j + "']", ExcelUtil.GetCellData(row, col + 8, sheetname), Property_type.XPath);
                    }
                    else
                    {
                        Console.WriteLine("Building Limit not given");
                    }
                    if (ExcelUtil.GetCellData(row, col + 9, sheetname) != "")
                    {
                        Perform.SelectDropDown(".//*[@id='cphMain_ctl_WorkflowManager_App_BOP_ctl_AppSection_BOP_ctl_LocationList_Repeater1_ctl_BOP_App_Location_" + i + "_ctl_BOP_App_BuildingList_" + i + "_Repeater1_" + i + "_ctl_BOP_App_Building_" + j + "_ddlPersonalPropertyLimitLossPayeeName_" + j + "']", ExcelUtil.GetCellData(row, col + 9, sheetname), Property_type.XPath);
                        Perform.SelectDropDown(".//*[@id='cphMain_ctl_WorkflowManager_App_BOP_ctl_AppSection_BOP_ctl_LocationList_Repeater1_ctl_BOP_App_Location_" + i + "_ctl_BOP_App_BuildingList_" + i + "_Repeater1_" + i + "ctl_BOP_App_Building_" + j + "_ddlPersonalPropertyLimitLossPayeeType_" + j + "']", ExcelUtil.GetCellData(row, col + 10, sheetname), Property_type.XPath);
                        Perform.SelectDropDown(".//*[@id='cphMain_ctl_WorkflowManager_App_BOP_ctl_AppSection_BOP_ctl_LocationList_Repeater1_ctl_BOP_App_Location_" + i + "_ctl_BOP_App_BuildingList_" + i + "_Repeater1_" + i + "_ctl_BOP_App_Building_" + j + "_ddlPersonalPropertyLimitATMA_" + j + "']", ExcelUtil.GetCellData(row, col + 11, sheetname), Property_type.XPath);
                    }
                    else
                    {
                        Console.WriteLine("Personal Property Limit not given");
                    }
                   
                        col = col + 12;
                    
                }
               
              
            }
        }
        public static void ContractorInfo(int row,string nofitems,string sheetname)
        {
            int col = 3;

            
        for(int i=0; i<Int32.Parse(nofitems); i++)
            {
                sheetname = "Contractor";
                if(i>0)
                {
                    Perform.Click(".//*[@id='cphMain_ctl_WorkflowManager_App_BOP_ctl_AppSection_BOP_ctl_ContractorsEQList_lnkBtnAdd']", Property_type.XPath);
                }
               
                Perform.EnterText(".//*[@id='cphMain_ctl_WorkflowManager_App_BOP_ctl_AppSection_BOP_ctl_ContractorsEQList_Repeater1_ctlScheduledItem_"+i+"_txtLimit_"+i+"']", ExcelUtil.GetCellData(row, col, sheetname), Property_type.XPath);
                Perform.EnterText(".//*[@id='cphMain_ctl_WorkflowManager_App_BOP_ctl_AppSection_BOP_ctl_ContractorsEQList_Repeater1_ctlScheduledItem_"+i+"_txtDescription_"+i+"']", ExcelUtil.GetCellData(row, col + 1, sheetname), Property_type.XPath);
                col = col + 6;
            }
            
        }
        public static void Additional_Insured(int row,string nofinsured,string sheetname)
        {
           
            int col = 1;
            for(int i=0;i<Int32.Parse(nofinsured);i++)
            {
                sheetname = "Additional Insured";
                Perform.Click(".//*[@id='cphMain_ctl_WorkflowManager_App_BOP_ctl_AppSection_BOP_ctl_App_AdditionalInsureds_DataGrid_additionalInsured']/tbody/tr[" + (i + 2 )+ "]/td[6]/input", Property_type.XPath);
                Perform.SelectDropDown(".//*[@id='cphMain_ctl_WorkflowManager_App_BOP_ctl_AppSection_BOP_ctl_App_AdditionalInsureds_DropDownList_additonalInsuredType']", ExcelUtil.GetCellData(row, col,sheetname), Property_type.XPath);
                Perform.EnterText(".//*[@id='cphMain_ctl_WorkflowManager_App_BOP_ctl_AppSection_BOP_ctl_App_AdditionalInsureds_TextBox_addInameOfOrg']", ExcelUtil.GetCellData(row, col + 1,sheetname), Property_type.XPath);
                Perform.Click(".//*[@id='cphMain_ctl_WorkflowManager_App_BOP_ctl_AppSection_BOP_ctl_App_AdditionalInsureds_Button_addInsSaveAdd']", Property_type.XPath);
                col = col + 7;
            }
            
        }
    }
}

