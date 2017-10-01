using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace VR_Personal_Auto
{
    class PageUtility
    {
        public static void AddDriverDetails(int row, string nofdriver, string Sheetname)
        {
            int columnum = 0;
            for (int i = 0; i < Int32.Parse(nofdriver); i++)
            {
                if (ExcelUtil.GetCellData(row, columnum, Sheetname) != "")
                {
                    Perform.Click("//*[@id='cphMain_ctl_Master_Edit_ctlDriverList_bnAddDriver']", Property_type.XPath);
                    Perform.EnterText(".//*[@id='cphMain_ctl_Master_Edit_ctlDriverList_Repeater1_ctlDriver_PPAControl_" + i + "_txtFirstName_" + i + "']", ExcelUtil.GetCellData(row, columnum, Sheetname), Property_type.XPath);
                    Perform.EnterText(".//*[@id='cphMain_ctl_Master_Edit_ctlDriverList_Repeater1_ctlDriver_PPAControl_" + i + "_txtMiddleName_" + i + "']", ExcelUtil.GetCellData(row, columnum + 1, Sheetname), Property_type.XPath);
                    Perform.EnterText(".//*[@id='cphMain_ctl_Master_Edit_ctlDriverList_Repeater1_ctlDriver_PPAControl_" + i + "_txtLastname_" + i + "']", ExcelUtil.GetCellData(row, columnum + 2, Sheetname), Property_type.XPath);
                    Perform.SelectDropDown(".//*[@id='cphMain_ctl_Master_Edit_ctlDriverList_Repeater1_ctlDriver_PPAControl_" + i + "_ddSuffix_" + i + "']", ExcelUtil.GetCellData(row, columnum + 3, Sheetname), Property_type.XPath);
                    Perform.EnterText(".//*[@id='cphMain_ctl_Master_Edit_ctlDriverList_Repeater1_ctlDriver_PPAControl_" + i + "_txtBirthDate_" + i + "']", ExcelUtil.GetCellData(row, columnum + 4, Sheetname), Property_type.XPath);

                    Perform.SelectDropDown(".//*[@id='cphMain_ctl_Master_Edit_ctlDriverList_Repeater1_ctlDriver_PPAControl_" + i + "_ddSex_" + i + "']", ExcelUtil.GetCellData(row, columnum + 5, Sheetname), Property_type.XPath);
                    Perform.SelectDropDown(".//*[@id='cphMain_ctl_Master_Edit_ctlDriverList_Repeater1_ctlDriver_PPAControl_" + i + "_ddMaritialStatus_" + i + "']", ExcelUtil.GetCellData(row, columnum + 6, Sheetname), Property_type.XPath);
                    Perform.EnterText(".//*[@id='cphMain_ctl_Master_Edit_ctlDriverList_Repeater1_ctlDriver_PPAControl_" + i + "_txtDLNumber_" + i + "']", ExcelUtil.GetCellData(row, columnum + 7, Sheetname), Property_type.XPath);

                    Perform.SelectDropDown(".//*[@id='cphMain_ctl_Master_Edit_ctlDriverList_Repeater1_ctlDriver_PPAControl_" + i + "_ddDLState_" + i + "']", ExcelUtil.GetCellData(row, columnum + 8, Sheetname), Property_type.XPath);
                    Perform.SelectDropDown(".//*[@id='cphMain_ctl_Master_Edit_ctlDriverList_Repeater1_ctlDriver_PPAControl_" + i + "_ddRelationToPolicyHolder_" + i + "']", ExcelUtil.GetCellData(row, columnum + 9, Sheetname), Property_type.XPath);
                    Perform.SelectDropDown(".//*[@id='cphMain_ctl_Master_Edit_ctlDriverList_Repeater1_ctlDriver_PPAControl_" + i + "_ddRatedOrExcludedDriver_" + i + "']", ExcelUtil.GetCellData(row, columnum + 10, Sheetname), Property_type.XPath);

                }
                columnum = columnum + 24;
            }
        }


        public static void AddVehicleDetails(int row, String nofdriver, String nofvehicle, string Sheetname)
        {
            int bodytypecol = 6;
            int ocdnum = 12;
            int colnum = 0;
            int garagenum = 28;
            for (int i = 0; i < Int32.Parse(nofvehicle); i++)
            {
                Perform.Click(".//*[@id='cphMain_ctl_Master_Edit_ctlVehicleList_btnAddvehicle']", Property_type.XPath);

                if (ExcelUtil.GetCellData(row, bodytypecol, Sheetname) == "CAR" || ExcelUtil.GetCellData(row, bodytypecol, Sheetname) == "PICKUP W/O CAMPER" || ExcelUtil.GetCellData(row, bodytypecol, Sheetname) == "SUV" || ExcelUtil.GetCellData(row, bodytypecol, Sheetname) == "VAN")
                {
                    Perform.EnterText(".//*[@id='cphMain_ctl_Master_Edit_ctlVehicleList_Repeater1_ctlVehicle_PPAControl_" + i + "_txtVinNumber_" + i + "']", ExcelUtil.GetCellData(row, colnum + 0, Sheetname), Property_type.XPath);
                    Perform.EnterText("//input[contains(@id,'txtYear_" + i + "')]", ExcelUtil.GetCellData(row, colnum + 1, Sheetname), Property_type.XPath);
                    Perform.EnterText(".//*[@id='cphMain_ctl_Master_Edit_ctlVehicleList_Repeater1_ctlVehicle_PPAControl_" + i + "_txtMake_" + i + "']", ExcelUtil.GetCellData(row, colnum + 2, Sheetname), Property_type.XPath);
                    Perform.EnterText(".//*[@id='cphMain_ctl_Master_Edit_ctlVehicleList_Repeater1_ctlVehicle_PPAControl_" + i + "_txtModel_" + i + "']", ExcelUtil.GetCellData(row, colnum + 3, Sheetname), Property_type.XPath);
                    Perform.SelectDropDown(".//*[@id='cphMain_ctl_Master_Edit_ctlVehicleList_Repeater1_ctlVehicle_PPAControl_" + i + "_ddBodyType_" + i + "']", ExcelUtil.GetCellData(row, colnum + 6, Sheetname), Property_type.XPath);
                    Perform.SelectDropDown(".//*[@id='cphMain_ctl_Master_Edit_ctlVehicleList_Repeater1_ctlVehicle_PPAControl_" + i + "_ddPrincipalDriver_" + i + "']", ExcelUtil.GetCellData(row, colnum + 11, Sheetname), Property_type.XPath);

                    for (int j = 0; j < Int32.Parse(nofdriver); j++)
                    {
                        if (ExcelUtil.GetCellData(row, ocdnum, Sheetname) != "")
                        {

                            Perform.SelectDropDown(".//*[@id='cphMain_ctl_Master_Edit_ctlVehicleList_Repeater1_ctlVehicle_PPAControl_" + j + "_ddOccDriver" + (j + 1) + "_" + j + "']", ExcelUtil.GetCellData(row, ocdnum, Sheetname), Property_type.XPath);
                        }
                        ocdnum = ocdnum + 1;

                    }
                    ocdnum = ocdnum + 24;

                }

                if (ExcelUtil.GetCellData(row, bodytypecol, Sheetname) == "MOTORCYCLE")
                {
                    Perform.EnterText(".//*[@id='cphMain_ctl_Master_Edit_ctlVehicleList_Repeater1_ctlVehicle_PPAControl_" + i + "_txtVinNumber_" + i + "']", ExcelUtil.GetCellData(row, colnum + 0, Sheetname), Property_type.XPath);
                    Perform.EnterText("//input[contains(@id,'txtYear_" + i + "')]", ExcelUtil.GetCellData(row, colnum + 1, Sheetname), Property_type.XPath);
                    Perform.EnterText(".//*[@id='cphMain_ctl_Master_Edit_ctlVehicleList_Repeater1_ctlVehicle_PPAControl_" + i + "_txtMake_" + i + "']", ExcelUtil.GetCellData(row, colnum + 2, Sheetname), Property_type.XPath);
                    Perform.EnterText(".//*[@id='cphMain_ctl_Master_Edit_ctlVehicleList_Repeater1_ctlVehicle_PPAControl_" + i + "_txtModel_" + i + "']", ExcelUtil.GetCellData(row, colnum + 3, Sheetname), Property_type.XPath);
                    Perform.SelectDropDown(".//*[@id='cphMain_ctl_Master_Edit_ctlVehicleList_Repeater1_ctlVehicle_PPAControl_" + i + "_ddBodyType_" + i + "']", ExcelUtil.GetCellData(row, colnum + 6, Sheetname), Property_type.XPath);
                    Perform.SelectDropDown(".//*[@id='cphMain_ctl_Master_Edit_ctlVehicleList_Repeater1_ctlVehicle_PPAControl_" + i + "_ddPrincipalDriver_" + i + "']", ExcelUtil.GetCellData(row, colnum + 11, Sheetname), Property_type.XPath);

                    Perform.EnterText("//input[contains(@id,'txtCostNew_1')]", ExcelUtil.GetCellData(row, colnum + 8, Sheetname), Property_type.XPath);
                    Perform.SelectDropDown("//select[contains(@id,'MotorCyleType_1')]", ExcelUtil.GetCellData(row, colnum + 9, Sheetname), Property_type.XPath);
                    Perform.EnterText(".//input[contains(@id,'_txtHorsePower_1')]", ExcelUtil.GetCellData(row, colnum + 10, Sheetname), Property_type.XPath);

                    for (int j = 0; j < Int32.Parse(nofdriver); j++)
                    {

                        if (ExcelUtil.GetCellData(row, ocdnum, Sheetname) != "")
                        {
                            Perform.SelectDropDown(".//*[@id='cphMain_ctl_Master_Edit_ctlVehicleList_Repeater1_ctlVehicle_PPAControl_" + j + "_ddOccDriver" + (j + 1) + "_" + j + "']", ExcelUtil.GetCellData(row, ocdnum, Sheetname), Property_type.XPath);
                            ocdnum = ocdnum + 1;
                        }

                    }
                    ocdnum = ocdnum + 24;
                }
                Perform.waitTillElementToAppear(".//*[@id='ui-id-" + garagenum + "']");
                Perform.Click(".//*[@id='ui-id-" + garagenum + "']", Property_type.XPath);
                Perform.waitTillElementToAppear(".//*[@id='cphMain_ctl_Master_Edit_ctlVehicleList_Repeater1_ctlVehicle_PPAControl_" + i + "_txtGaragedStreetNum_" + i + "']");
                Perform.EnterText(".//*[@id='cphMain_ctl_Master_Edit_ctlVehicleList_Repeater1_ctlVehicle_PPAControl_" + i + "_txtGaragedStreetNum_" + i + "']", ExcelUtil.GetCellData(row, colnum + 17, Sheetname), Property_type.XPath);
                Perform.EnterText(".//*[@id='cphMain_ctl_Master_Edit_ctlVehicleList_Repeater1_ctlVehicle_PPAControl_" + i + "_txtGaragedStreet_" + i + "']", ExcelUtil.GetCellData(row, colnum + 18, Sheetname), Property_type.XPath);
                Perform.EnterText(".//*[@id='cphMain_ctl_Master_Edit_ctlVehicleList_Repeater1_ctlVehicle_PPAControl_" + i + "_txtGaragedStreet_" + i + "']", ExcelUtil.GetCellData(row, colnum + 19, Sheetname), Property_type.XPath);
                Perform.EnterText(".//*[@id='cphMain_ctl_Master_Edit_ctlVehicleList_Repeater1_ctlVehicle_PPAControl_" + i + "_txtGaragedCity_" + i + "']", ExcelUtil.GetCellData(row, colnum + 20, Sheetname), Property_type.XPath);
                Perform.SelectDropDown(".//*[@id='cphMain_ctl_Master_Edit_ctlVehicleList_Repeater1_ctlVehicle_PPAControl_" + i + "_ddGaragedState_" + i + "']", ExcelUtil.GetCellData(row, colnum + 21, Sheetname), Property_type.XPath);
                Perform.EnterText(".//*[@id='cphMain_ctl_Master_Edit_ctlVehicleList_Repeater1_ctlVehicle_PPAControl_" + i + "_txtGaragedZip_" + i + "']", ExcelUtil.GetCellData(row, colnum + 22, Sheetname), Property_type.XPath);
                Perform.EnterText(".//*[@id='cphMain_ctl_Master_Edit_ctlVehicleList_Repeater1_ctlVehicle_PPAControl_" + i + "_txtGaragedCounty_" + i + "']", ExcelUtil.GetCellData(row, colnum + 23, Sheetname), Property_type.XPath);
                bodytypecol = bodytypecol + 24;
                garagenum = garagenum + 14;
                colnum = colnum + 24;

            }



        }
        public static void Coveragevalue(int row, string nofvehicle, string Sheetname)
        {
            int bodytypecol = 6;
            int policycol = 11;



            Perform.SelectDropDown(".//*[@id='cphMain_ctl_Master_Edit_ctlCoverage_PPA_ddLiabType']", ExcelUtil.GetCellData(row, 0, Sheetname), Property_type.XPath);
            Perform.SelectDropDown(".//*[@id='cphMain_ctl_Master_Edit_ctlCoverage_PPA_ddBodilyInjury']", ExcelUtil.GetCellData(row, 1, Sheetname), Property_type.XPath);

            Perform.SelectDropDown(".//*[@id='cphMain_ctl_Master_Edit_ctlCoverage_PPA_ddPropertyDamage']", ExcelUtil.GetCellData(row, 2, Sheetname), Property_type.XPath);
            Perform.SelectDropDown(".//*[@id='cphMain_ctl_Master_Edit_ctlCoverage_PPA_ddmedicalPayments']", ExcelUtil.GetCellData(row, 3, Sheetname), Property_type.XPath);
            Perform.SelectDropDown(".//*[@id='cphMain_ctl_Master_Edit_ctlCoverage_PPA_ddUmUmiBi']", ExcelUtil.GetCellData(row, 4, Sheetname), Property_type.XPath);

            Perform.SelectDropDown(".//*[@id='cphMain_ctl_Master_Edit_ctlCoverage_PPA_ddUmPd']", ExcelUtil.GetCellData(row, 5, Sheetname), Property_type.XPath);
            Perform.SelectDropDown(".//*[@id='cphMain_ctl_Master_Edit_ctlCoverage_PPA_ddUmPdDeductible']", ExcelUtil.GetCellData(row, 6, Sheetname), Property_type.XPath);

            for (int i = 0; i < Int32.Parse(nofvehicle); i++)
            {   
                if (i>=1)
                {
                    Console.WriteLine(i+" Coverage");
                    Perform.Click(".//*[@id='cphMain_ctl_Master_Edit_ctlCoverage_PPA_ctlCoverage_PPA_Vehicle_List_Repeater1_ctlCoverage_PPA_VehicleSpecific_"+i+"_lblAccordHeader_"+i+"']", Property_type.XPath);
                }
                if (ExcelUtil.GetCellData(row, bodytypecol, "Vehicle") == "CAR" || ExcelUtil.GetCellData(row, bodytypecol, "Vehicle") == "PICKUP W/O CAMPER" || ExcelUtil.GetCellData(row, bodytypecol, "Vehicle") == "SUV" || ExcelUtil.GetCellData(row, bodytypecol, "Vehicle") == "VAN")
                {
                    if (ExcelUtil.GetCellData(row, policycol, Sheetname) == "FULL COVERAGE")
                    {
                        Perform.SelectDropDown(".//*[@id='ddPolicy"+i+"']", ExcelUtil.GetCellData(row, policycol, Sheetname), Property_type.XPath);
                        Perform.SelectDropDown(".//*[@id='cphMain_ctl_Master_Edit_ctlCoverage_PPA_ctlCoverage_PPA_Vehicle_List_Repeater1_ctlCoverage_PPA_VehicleSpecific_" + i + "_ddComprehensive_" + i + "']", ExcelUtil.GetCellData(row, policycol + 1, Sheetname), Property_type.XPath);
                        Perform.SelectDropDown(".//*[@id='cphMain_ctl_Master_Edit_ctlCoverage_PPA_ctlCoverage_PPA_Vehicle_List_Repeater1_ctlCoverage_PPA_VehicleSpecific_" + i + "_ddCollision_" + i + "']", ExcelUtil.GetCellData(row, policycol + 2, Sheetname), Property_type.XPath);
                        Perform.SelectDropDown(".//*[@id='cphMain_ctl_Master_Edit_ctlCoverage_PPA_ctlCoverage_PPA_Vehicle_List_Repeater1_ctlCoverage_PPA_VehicleSpecific_" + i + "_ddTowing_" + i + "']", ExcelUtil.GetCellData(row, policycol + 3, Sheetname), Property_type.XPath);
                        Perform.SelectDropDown(".//*[@id='cphMain_ctl_Master_Edit_ctlCoverage_PPA_ctlCoverage_PPA_Vehicle_List_Repeater1_ctlCoverage_PPA_VehicleSpecific_" + i + "_ddTransportation_" + i + "']", ExcelUtil.GetCellData(row, policycol + 4, Sheetname), Property_type.XPath);
                        Perform.SelectDropDown(".//*[@id='cphMain_ctl_Master_Edit_ctlCoverage_PPA_ctlCoverage_PPA_Vehicle_List_Repeater1_ctlCoverage_PPA_VehicleSpecific_" + i + "_ddRadio_" + i + "']", ExcelUtil.GetCellData(row, policycol + 5, Sheetname), Property_type.XPath);

                        Perform.SelectDropDown(".//*[@id='cphMain_ctl_Master_Edit_ctlCoverage_PPA_ctlCoverage_PPA_Vehicle_List_Repeater1_ctlCoverage_PPA_VehicleSpecific_" + i + "_ddAudioVisual_" + i + "']", ExcelUtil.GetCellData(row, policycol + 6, Sheetname), Property_type.XPath);
                        Perform.SelectDropDown(".//*[@id='cphMain_ctl_Master_Edit_ctlCoverage_PPA_ctlCoverage_PPA_Vehicle_List_Repeater1_ctlCoverage_PPA_VehicleSpecific_" + i + "_ddMedia_" + i + "']", ExcelUtil.GetCellData(row, policycol + 7, Sheetname), Property_type.XPath);
                        if(ExcelUtil.GetCellData(row, policycol+9, Sheetname) =="YES")
                            {
                            if(Property_Collection.driver.FindElement(By.XPath(".//*[@id='cphMain_ctl_Master_Edit_ctlCoverage_PPA_ctlCoverage_PPA_Vehicle_List_Repeater1_ctlCoverage_PPA_VehicleSpecific_"+i+"_chkAutoLoanLease_"+i+"']")).Displayed)
                            {
                                Perform.Click(".//*[@id='cphMain_ctl_Master_Edit_ctlCoverage_PPA_ctlCoverage_PPA_Vehicle_List_Repeater1_ctlCoverage_PPA_VehicleSpecific_" + i + "_chkAutoLoanLease_" + i + "']", Property_type.XPath);
                            }
                        }
                    }
                    if (ExcelUtil.GetCellData(row, policycol, Sheetname) == "LIABILITY ONLY")
                    {
                        Perform.SelectDropDown(".//*[@id='ddPolicy"+i+"']", ExcelUtil.GetCellData(row, policycol, Sheetname), Property_type.XPath);
                    }
                }

                if(ExcelUtil.GetCellData(row, bodytypecol, "Vehicle") == "MOTORCYCLE")
                {
                    if (ExcelUtil.GetCellData(row, policycol, Sheetname) == "FULL COVERAGE")
                    {
                        Console.WriteLine("Motor Info");
                        Perform.SelectDropDown(".//*[@id='cphMain_ctl_Master_Edit_ctlCoverage_PPA_ctlCoverage_PPA_Vehicle_List_Repeater1_ctlCoverage_PPA_VehicleSpecific_" + i + "_ddComprehensive_" + i + "']", ExcelUtil.GetCellData(row, policycol + 1, Sheetname), Property_type.XPath);
                        Perform.SelectDropDown(".//*[@id='cphMain_ctl_Master_Edit_ctlCoverage_PPA_ctlCoverage_PPA_Vehicle_List_Repeater1_ctlCoverage_PPA_VehicleSpecific_" + i + "_ddCollision_" + i + "']", ExcelUtil.GetCellData(row, policycol + 2, Sheetname), Property_type.XPath);
                        Perform.SelectDropDown(".//*[@id='cphMain_ctl_Master_Edit_ctlCoverage_PPA_ctlCoverage_PPA_Vehicle_List_Repeater1_ctlCoverage_PPA_VehicleSpecific_" + i + "_ddTowing_" + i + "']", ExcelUtil.GetCellData(row, policycol + 3, Sheetname), Property_type.XPath);
                        Perform.EnterText(".//*[@id='cphMain_ctl_Master_Edit_ctlCoverage_PPA_ctlCoverage_PPA_Vehicle_List_Repeater1_ctlCoverage_PPA_VehicleSpecific_" + i + "_txtMotorEquip_" + i + "']", ExcelUtil.GetCellData(row, policycol + 8, Sheetname), Property_type.XPath);
                        Perform.SelectDropDown(".//*[@id='cphMain_ctl_Master_Edit_ctlCoverage_PPA_ctlCoverage_PPA_Vehicle_List_Repeater1_ctlCoverage_PPA_VehicleSpecific_" + i + "_ddMedia_" + i + "']", ExcelUtil.GetCellData(row, policycol + 7 , Sheetname), Property_type.XPath);

                    }
                    if (ExcelUtil.GetCellData(row, policycol, Sheetname) == "LIABILITY ONLY")
                    {
                        Perform.SelectDropDown(".//*[@id='ddPolicy"+i+"']", ExcelUtil.GetCellData(row, policycol, Sheetname), Property_type.XPath);
                    }
                }
                    policycol = policycol + 10;
                    bodytypecol = bodytypecol + 24;
                
                
            }


        }
    }

}










