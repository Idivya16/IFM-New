using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DIA_HOM_NB
{
    class PageUtility
    {
        public static void Updates(int row, string sheetname)
        {
            sheetname = "Updates";
            Perform.waitTillElementToAppear("//a[@id='P_L_V_v39w9_t17_c0w0_PC_T1T']/span");
            Perform.Click("//a[@id='P_L_V_v39w9_t17_c0w0_PC_T1T']/span");
            Perform.waitTillElementToAppear(".//*[@id='P_L_V_v39w9_t17_c0w0_PC_t1i0_RoofInfoYearUpdatedInsNumeric_RoofInfoYearUpdatedInsNumeric']");
            //Roof Information
            if (ExcelUtil.GetCellData(row, 1, sheetname) != "")
            {
                Perform.EnterText(".//*[@id='P_L_V_v39w9_t17_c0w0_PC_t1i0_RoofInfoYearUpdatedInsNumeric_RoofInfoYearUpdatedInsNumeric']", ExcelUtil.GetCellData(row, 1, sheetname));

                if (ExcelUtil.GetCellData(row, 2, sheetname) == "Partial")
                {
                    Perform.Click(".//*[@id='P_L_V_v39w9_t17_c0w0_PC_t1i0_RoofInfoPartialRadioButton']");
                }
                else
                {
                    Perform.Click(".//*[@id='P_L_V_v39w9_t17_c0w0_PC_t1i0_RoofInfoCompleteRadioButton']");
                }
                if (ExcelUtil.GetCellData(row, 3, sheetname) != "" && ExcelUtil.GetCellData(row, 3, sheetname) == "YES")
                {
                    Perform.Click(".//*[@id='P_L_V_v39w9_t17_c0w0_PC_t1i0_RoofInfoAsphaltShingleInsCheckbox']");
                }
                if (ExcelUtil.GetCellData(row, 4, sheetname) != "" && ExcelUtil.GetCellData(row, 4, sheetname) == "YES")
                {
                    Perform.Click(".//*[@id='P_L_V_v39w9_t17_c0w0_PC_t1i0_RoofInfoWoodInsCheckbox']");
                }
                if (ExcelUtil.GetCellData(row, 5, sheetname) != "" && ExcelUtil.GetCellData(row, 5, sheetname) == "YES")
                {
                    Perform.Click(".//*[@id='P_L_V_v39w9_t17_c0w0_PC_t1i0_RoofInfoOtherInsCheckbox']");
                }
                if (ExcelUtil.GetCellData(row, 6, sheetname) != "" && ExcelUtil.GetCellData(row, 6, sheetname) == "YES")
                {
                    Perform.Click(".//*[@id='P_L_V_v39w9_t17_c0w0_PC_t1i0_RoofInfoSlateInsCheckbox']");
                }
                if (ExcelUtil.GetCellData(row, 7, sheetname) != "" && ExcelUtil.GetCellData(row, 7, sheetname) == "YES")
                {
                    Perform.Click(".//*[@id='P_L_V_v39w9_t17_c0w0_PC_t1i0_RoofInfoMetalInsCheckbox']");
                }
            }
            //Central Heat Information
            if (ExcelUtil.GetCellData(row, 8, sheetname) != "")
            {
                Perform.EnterText(".//*[@id='P_L_V_v39w9_t17_c0w0_PC_t1i0_CentralHeatInfoYearUpdatedInsNumeric_CentralHeatInfoYearUpdatedInsNumeric']", ExcelUtil.GetCellData(row, 1, sheetname));

                if (ExcelUtil.GetCellData(row, 9, sheetname) == "Partial")
                {
                    Perform.Click(".//*[@id='P_L_V_v39w9_t17_c0w0_PC_t1i0_CentralHeatInfoPartialRadioButton']");
                }
                else
                {
                    Perform.Click(".//*[@id='P_L_V_v39w9_t17_c0w0_PC_t1i0_CentralHeatInfoCompleteRadioButton']");
                }
                if (ExcelUtil.GetCellData(row, 10, sheetname) != "" && ExcelUtil.GetCellData(row, 10, sheetname) == "YES")
                {
                    Perform.Click(".//*[@id='P_L_V_v39w9_t17_c0w0_PC_t1i0_CentralHeatInfoOilInsCheckbox']");
                }
                if (ExcelUtil.GetCellData(row, 11, sheetname) != "" && ExcelUtil.GetCellData(row, 11, sheetname) == "YES")
                {
                    Perform.Click(".//*[@id='P_L_V_v39w9_t17_c0w0_PC_t1i0_CentralHeatInfoGasInsCheckbox']");
                }
                if (ExcelUtil.GetCellData(row, 12, sheetname) != "" && ExcelUtil.GetCellData(row, 12, sheetname) == "YES")
                {
                    Perform.Click(".//*[@id='P_L_V_v39w9_t17_c0w0_PC_t1i0_CentralHeatInfoElectricInsCheckbox']");
                }
                if (ExcelUtil.GetCellData(row, 13, sheetname) != "" && ExcelUtil.GetCellData(row, 13, sheetname) == "YES")
                {
                    Perform.Click(".//*[@id='P_L_V_v39w9_t17_c0w0_PC_t1i0_CentralHeatInfoOtherInsCheckbox']");
                }
            }
            //Supplemental Heat Information
            if (ExcelUtil.GetCellData(row, 14, sheetname) != "")
            {
                Perform.EnterText(".//*[@id='P_L_V_v39w9_t17_c0w0_PC_t1i0_SupplementalHeatInfoYearUpdatedInsNumeric_SupplementalHeatInfoYearUpdatedInsNumeric']", ExcelUtil.GetCellData(row, 1, sheetname));

                if (ExcelUtil.GetCellData(row, 15, sheetname) == "Partial")
                {
                    Perform.Click(".//*[@id='P_L_V_v39w9_t17_c0w0_PC_t1i0_SupplementalHeatInfoPartialRadioButton']");
                }
                else
                {
                    Perform.Click(".//*[@id='P_L_V_v39w9_t17_c0w0_PC_t1i0_SupplementalHeatInfoCompleteRadioButton']");
                }
                if (ExcelUtil.GetCellData(row, 16, sheetname) != "" && ExcelUtil.GetCellData(row, 16, sheetname) == "YES")
                {
                    Perform.Click(".//*[@id='P_L_V_v39w9_t17_c0w0_PC_t1i0_SupplementalHeatInfoNotApplicableInsCheckbox']");
                }
                if (ExcelUtil.GetCellData(row, 17, sheetname) != "" && ExcelUtil.GetCellData(row, 17, sheetname) == "YES")
                {
                    Perform.Click(".//*[@id='P_L_V_v39w9_t17_c0w0_PC_t1i0_SupplementalHeatInfoFireplaceInsCheckbox']");
                }
                if (ExcelUtil.GetCellData(row, 18, sheetname) != "" && ExcelUtil.GetCellData(row, 18, sheetname) == "YES")
                {
                    Perform.Click(".//*[@id='P_L_V_v39w9_t17_c0w0_PC_t1i0_SupplementalHeatInfoFireplaceInsertInsCheckbox']");
                }
                if (ExcelUtil.GetCellData(row, 19, sheetname) != "" && ExcelUtil.GetCellData(row, 19, sheetname) == "YES")
                {
                    Perform.Click(".//*[@id='P_L_V_v39w9_t17_c0w0_PC_t1i0_SupplementalHeatInfoSolidFuelInsCheckbox']");
                }
                if (ExcelUtil.GetCellData(row, 20, sheetname) != "" && ExcelUtil.GetCellData(row, 20, sheetname) == "YES")
                {
                    Perform.Click(".//*[@id='P_L_V_v39w9_t17_c0w0_PC_t1i0_SupplementalHeatInfoBurningUnitInsCheckbox']");
                }
                if (ExcelUtil.GetCellData(row, 21, sheetname) != "" && ExcelUtil.GetCellData(row, 21, sheetname) == "YES")
                {
                    Perform.Click(".//*[@id='P_L_V_v39w9_t17_c0w0_PC_t1i0_SupplementalHeatInfoSpaceHeaterInsCheckbox']");
                }
            }
            //Electric Service Information
            if (ExcelUtil.GetCellData(row, 22, sheetname) != "")
            {
                Perform.EnterText(".//*[@id='P_L_V_v39w9_t17_c0w0_PC_t1i0_ElectricServiceInfoYearUpdatedInsNumeric_ElectricServiceInfoYearUpdatedInsNumeric']", ExcelUtil.GetCellData(row, 1, sheetname));

                if (ExcelUtil.GetCellData(row, 23, sheetname) == "Partial")
                {
                    Perform.Click(".//*[@id='P_L_V_v39w9_t17_c0w0_PC_t1i0_ElectricServiceInfoPartialRadioButton']");
                }
                else
                {
                    Perform.Click(".//*[@id='P_L_V_v39w9_t17_c0w0_PC_t1i0_ElectricServiceInfoCompleteRadioButton']");
                }
                if (ExcelUtil.GetCellData(row, 24, sheetname) != "" && ExcelUtil.GetCellData(row, 24, sheetname) == "YES")
                {
                    Perform.Click(".//*[@id='P_L_V_v39w9_t17_c0w0_PC_t1i0_ElectricServiceInfoCircuitBreakerInsCheckbox']");
                }
                if (ExcelUtil.GetCellData(row, 25, sheetname) != "" && ExcelUtil.GetCellData(row, 25, sheetname) == "YES")
                {
                    Perform.Click(".//*[@id='P_L_V_v39w9_t17_c0w0_PC_t1i0_ElectricServiceInfoFusesInsCheckbox']");
                }
                if (ExcelUtil.GetCellData(row, 26, sheetname) != "" && ExcelUtil.GetCellData(row, 26, sheetname) == "YES")
                {
                    Perform.Click(".//*[@id='P_L_V_v39w9_t17_c0w0_PC_t1i0_ElectricServiceInfo60AmperageInsCheckbox']");
                }
                if (ExcelUtil.GetCellData(row, 27, sheetname) != "" && ExcelUtil.GetCellData(row, 27, sheetname) == "YES")
                {
                    Perform.Click(".//*[@id='P_L_V_v39w9_t17_c0w0_PC_t1i0_ElectricServiceInfo100AmperageInsCheckbox']");
                }
                if (ExcelUtil.GetCellData(row, 28, sheetname) != "" && ExcelUtil.GetCellData(row, 28, sheetname) == "YES")
                {
                    Perform.Click(".//*[@id='P_L_V_v39w9_t17_c0w0_PC_t1i0_ElectricServiceInfo200AmperageInsCheckbox']");
                }
            }
            //Plumbing Information
            if (ExcelUtil.GetCellData(row, 29, sheetname) != "")
            {
                Perform.EnterText(".//*[@id='P_L_V_v39w9_t17_c0w0_PC_t1i0_PlumbingInfoYearUpdatedInsNumeric_PlumbingInfoYearUpdatedInsNumeric']", ExcelUtil.GetCellData(row, 29, sheetname));

                if (ExcelUtil.GetCellData(row, 30, sheetname) == "Partial")
                {
                    Perform.Click(".//*[@id='P_L_V_v39w9_t17_c0w0_PC_t1i0_PlumbingInfoPartialRadioButton']");
                }
                else
                {
                    Perform.Click(".//*[@id='P_L_V_v39w9_t17_c0w0_PC_t1i0_PlumbingInfoCompleteRadioButton']");
                }
                if (ExcelUtil.GetCellData(row, 31, sheetname) != "" && ExcelUtil.GetCellData(row, 31, sheetname) == "YES")
                {
                    Perform.Click(".//*[@id='P_L_V_v39w9_t17_c0w0_PC_t1i0_PlumbingInfoPlasticInsCheckbox']");
                }
                if (ExcelUtil.GetCellData(row, 32, sheetname) != "" && ExcelUtil.GetCellData(row, 32, sheetname) == "YES")
                {
                    Perform.Click(".//*[@id='P_L_V_v39w9_t17_c0w0_PC_t1i0_PlumbingInfoGalvanizedInsCheckbox']");
                }
                if (ExcelUtil.GetCellData(row, 33, sheetname) != "" && ExcelUtil.GetCellData(row, 33, sheetname) == "YES")
                {
                    Perform.Click(".//*[@id='P_L_V_v39w9_t17_c0w0_PC_t1i0_PlumbingInfoCopperInsCheckbox']");
                }
                Perform.EnterText(".//*[@id='P_L_V_v39w9_t17_c0w0_PC_t1i0_DescriptionCommentsInsTextBox']", ExcelUtil.GetCellData(row, 34, sheetname));
            }
            //Windows
            if (ExcelUtil.GetCellData(row, 35, sheetname) != "")
            {
                Perform.EnterText(".//*[@id='P_L_V_v39w9_t17_c0w0_PC_t1i0_WindowsYearUpdatedInsNumeric_WindowsYearUpdatedInsNumeric']", ExcelUtil.GetCellData(row, 35, sheetname));

                if (ExcelUtil.GetCellData(row, 36, sheetname) == "Partial")
                {
                    Perform.Click(".//*[@id='P_L_V_v39w9_t17_c0w0_PC_t1i0_WindowsPartialRadioButton']");
                }
                else
                {
                    Perform.Click(".//*[@id='P_L_V_v39w9_t17_c0w0_PC_t1i0_WindowsCompleteRadioButton']");
                }
            }
            //Inspection
            if (ExcelUtil.GetCellData(row, 38, sheetname) != "")
            {
                Perform.EnterText(".//*[@id='P_L_V_v39w9_t17_c0w0_PC_t1i0_InspectionRemarksInsTextBox']", ExcelUtil.GetCellData(row, 37, sheetname));
                Perform.EnterTextFocus(".//*[@id='P_L_V_v39w9_t17_c0w0_PC_t1i0_InspectionInsDateTime_InspectionInsDateTime']", ExcelUtil.GetCellData(row, 38, sheetname));
                if (ExcelUtil.GetCellData(row, 39, sheetname) == "Partial")
                {
                    Perform.Click(".//*[@id='P_L_V_v39w9_t17_c0w0_PC_t1i0_InspectionPartialRadioButton']");
                }
                else
                {
                    Perform.Click(".//*[@id='P_L_V_v39w9_t17_c0w0_PC_t1i0_InspectionCompleteRadioButton']");
                }
            }

        }
        public static void InlandMarine(int row, string sheetname)
        {
            sheetname = "Inland_Marine";
            int col = 2;
            string nofinland = ExcelUtil.GetCellData(row, 0, sheetname);
            for (int i = 0; i < Int32.Parse(nofinland); i++)
            {
                System.Threading.Thread.Sleep(500);
                Perform.Click(".//*[@id='AddInlandMarineToolStripButtonMiddle']/a");
                Perform.EnterTextFocus(".//*[@id='P_L_V_v39w9_t19_c0w0_PC_t0i"+i+"_InsCoverageControl_CoverageInsCombo_D_I']", ExcelUtil.GetCellData(row, col, sheetname));
                if (ExcelUtil.GetCellData(row, col, sheetname) == "Jewelry")
                {
                    Perform.EnterText(".//*[@id='P_L_V_v39w9_t19_c0w0_PC_t0i" + i + "_InsCoverageControl_CoverageControlASPxCallbackPanel_A_12179_70089_A_12179_70089_MainLimitLimit_A_12179_70089_MainLimitLimit_I']", ExcelUtil.GetCellData(row, col + 1, sheetname));
                    Perform.SelectTextDropDown(".//*[@id='P_L_V_v39w9_t19_c0w0_PC_t0i" + i + "_InsCoverageControl_CoverageControlASPxCallbackPanel_A_12179_70089_A_12179_70089_Deductible1Limit_D_I']", ExcelUtil.GetCellData(row, col + 2, sheetname));
                }
                if (ExcelUtil.GetCellData(row, col, sheetname) == "Silverware/Goldware")
                {
                    Perform.EnterText(".//*[@id='P_L_V_v39w9_t19_c0w0_PC_t0i"+i+"_InsCoverageControl_CoverageControlASPxCallbackPanel_A_12180_70090_A_12180_70090_MainLimitLimit_A_12180_70090_MainLimitLimit_I']", ExcelUtil.GetCellData(row, col + 1, sheetname));
                    Perform.SelectTextDropDown(".//*[@id='P_L_V_v39w9_t19_c0w0_PC_t0i"+i+"_InsCoverageControl_CoverageControlASPxCallbackPanel_A_12180_70090_A_12180_70090_Deductible1Limit_D_I']", ExcelUtil.GetCellData(row, col + 2, sheetname));
                }
                if (ExcelUtil.GetCellData(row, col, sheetname) == "Fine Arts with Breakage")
                {
                    Perform.EnterText(".//*[@id='P_L_V_v39w9_t19_c0w0_PC_t0i"+i+"_InsCoverageControl_CoverageControlASPxCallbackPanel_A_12174_70084_A_12174_70084_MainLimitLimit_A_12174_70084_MainLimitLimit_I']", ExcelUtil.GetCellData(row, col + 1, sheetname));
                    Perform.SelectTextDropDown(".//*[@id='P_L_V_v39w9_t19_c0w0_PC_t0i"+i+"_InsCoverageControl_CoverageControlASPxCallbackPanel_A_12174_70084_A_12174_70084_Deductible1Limit_D_I']", ExcelUtil.GetCellData(row, col + 2, sheetname));
                }
                if (ExcelUtil.GetCellData(row, col, sheetname) == "Musical Instruments Non-Professional")
                {
                    Perform.EnterText(".//*[@id='P_L_V_v39w9_t19_c0w0_PC_t0i"+i+"_InsCoverageControl_CoverageControlASPxCallbackPanel_A_12182_70094_A_12182_70094_MainLimitLimit_A_12182_70094_MainLimitLimit_I']", ExcelUtil.GetCellData(row, col + 1, sheetname));
                    Perform.SelectTextDropDown(".//*[@id='P_L_V_v39w9_t19_c0w0_PC_t0i"+i+"_InsCoverageControl_CoverageControlASPxCallbackPanel_A_12182_70094_A_12182_70094_Deductible1Limit_D_I']", ExcelUtil.GetCellData(row, col + 2, sheetname));
                }
                if (ExcelUtil.GetCellData(row, col, sheetname) == "Fine Arts without Breakage")
                {
                    Perform.EnterText(".//*[@id='P_L_V_v39w9_t19_c0w0_PC_t0i"+i+"_InsCoverageControl_CoverageControlASPxCallbackPanel_A_12175_70085_A_12175_70085_MainLimitLimit_A_12175_70085_MainLimitLimit_I']", ExcelUtil.GetCellData(row, col + 1, sheetname));
                    Perform.SelectTextDropDown(".//*[@id='P_L_V_v39w9_t19_c0w0_PC_t0i"+i+"_InsCoverageControl_CoverageControlASPxCallbackPanel_A_12175_70085_A_12175_70085_Deductible1Limit_D_I']", ExcelUtil.GetCellData(row, col + 2, sheetname));
                }
                if (ExcelUtil.GetCellData(row, col, sheetname) == "Bicycles")
                {
                    Perform.EnterText(".//*[@id='P_L_V_v39w9_t19_c0w0_PC_t0i"+i+"_InsCoverageControl_CoverageControlASPxCallbackPanel_A_12171_70077_A_12171_70077_MainLimitLimit_A_12171_70077_MainLimitLimit_I']", ExcelUtil.GetCellData(row, col + 1, sheetname));
                    Perform.SelectTextDropDown(".//*[@id='P_L_V_v39w9_t19_c0w0_PC_t0i"+i+"_InsCoverageControl_CoverageControlASPxCallbackPanel_A_12171_70077_A_12171_70077_Deductible1Limit_D_I']", ExcelUtil.GetCellData(row, col + 2, sheetname));
                }
                Perform.EnterText(".//*[@id='P_L_V_v39w9_t19_c0w0_PC_t0i"+i+"_InlandMarineDescriptionInsTextBox']", ExcelUtil.GetCellData(row, col + 3, sheetname));
                if (ExcelUtil.GetCellData(row, col + 4, sheetname) != "")
                {
                    Perform.EnterText(".//*[@id='P_L_V_v39w9_t19_c0w0_PC_t0i"+i+"_StorageLocationInsTextBox']", ExcelUtil.GetCellData(row, col + 4, sheetname));
                }
                if (ExcelUtil.GetCellData(row, col + 5, sheetname) != "")
                {
                    Perform.Click(".//*[@id='P_L_V_v39w9_t19_c0w0_PC_t0i"+i+"_StatedAmountInsCheckBox']");
                }
                if (ExcelUtil.GetCellData(row, col + 6, sheetname) != "")
                {
                    Perform.SelectTextDropDown(".//*[@id='P_L_V_v39w9_t19_c0w0_PC_t0i"+i+"_TypeInsCombo_D_I']", ExcelUtil.GetCellData(row, col + 6, sheetname));
                }
                if (ExcelUtil.GetCellData(row, col + 7, sheetname) != "")
                {
                    Perform.EnterText(".//*[@id='P_L_V_v39w9_t19_c0w0_PC_t0i"+i+"_YearInsNumeric_YearInsNumeric']", ExcelUtil.GetCellData(row, col + 7, sheetname));
                }
                if (ExcelUtil.GetCellData(row, col + 8, sheetname) != "")
                {
                    Perform.EnterText(".//*[@id='P_L_V_v39w9_t19_c0w0_PC_t0i"+i+"_MakeBrandInsTextBox']", ExcelUtil.GetCellData(row, col + 7, sheetname));
                }
                if (ExcelUtil.GetCellData(row, col + 9, sheetname) != "")
                {
                    Perform.EnterText(".//*[@id='P_L_V_v39w9_t19_c0w0_PC_t0i"+i+"_ModelNameInsTextBox']", ExcelUtil.GetCellData(row, col + 9, sheetname));
                }
                if (ExcelUtil.GetCellData(row, col + 10, sheetname) != "")
                {
                    Perform.EnterText(".//*[@id='P_L_V_v39w9_t19_c0w0_PC_t0i"+i+"_ArtistNameInsTextBox']", ExcelUtil.GetCellData(row, col + 10, sheetname));
                }
                if (ExcelUtil.GetCellData(row, col + 11, sheetname) != "")
                {
                    Perform.EnterText(".//*[@id='P_L_V_v39w9_t19_c0w0_PC_t0i"+i+"_SerialNumberInsTextBox']", ExcelUtil.GetCellData(row, col + 11, sheetname));
                }
                Perform.Click(".//*[@id='SaveToolStripButtonMiddle']/a");
                col = col + 13;
            }
        }
            public static void RVWatercraft(int row,string sheetname)
        {
            sheetname = "R_V_Watercraft";
            string nofwatercraft= ExcelUtil.GetCellData(row, 0, sheetname);
            int col = 2;
            for(int i = 0;i<Int32.Parse(nofwatercraft);i++)
            {
                Perform.Click(".//*[@id='AddToolStripButtonMiddle']/a");
                Perform.SelectTextDropDown(".//*[@id='P_L_V_v39w9_t20_c0w0_PC_t0i"+i+"_TypeWatercraftInsCombo_D_I']", ExcelUtil.GetCellData(row, col, sheetname));
                if (ExcelUtil.GetCellData(row, col + 1, sheetname) != "")
                {
                    Perform.EnterText(".//*[@id='P_L_V_v39w9_t20_c0w0_PC_t0i"+i+"_YearInsNumeric_YearInsNumeric']", ExcelUtil.GetCellData(row, col + 1, sheetname));
                }
                if (ExcelUtil.GetCellData(row, col + 2, sheetname) != "")
                {
                    Perform.EnterText(".//*[@id='P_L_V_v39w9_t20_c0w0_PC_t0i"+i+"_ManufacturerInsTextBox']", ExcelUtil.GetCellData(row, col + 2, sheetname));
                }
                if (ExcelUtil.GetCellData(row, col + 3, sheetname) != "")
                {
                    Perform.EnterText(".//*[@id='P_L_V_v39w9_t20_c0w0_PC_t0i"+i+"_ModelInsTextBox']", ExcelUtil.GetCellData(row, col + 3, sheetname));
                }
                if (ExcelUtil.GetCellData(row, col + 4, sheetname) != "")
                {
                    Perform.EnterText(".//*[@id='P_L_V_v39w9_t20_c0w0_PC_t0i"+i+"_SerialNumberInsTextBox']", ExcelUtil.GetCellData(row, col + 4, sheetname));
                }
                if (ExcelUtil.GetCellData(row, col + 5, sheetname) != "")
                {
                    Perform.EnterText(".//*[@id='P_L_V_v39w9_t20_c0w0_PC_t0i"+i+"_HorsepowerRvWatercraftInsNumeric_HorsepowerRvWatercraftInsNumeric']", ExcelUtil.GetCellData(row, col + 5, sheetname));
                }
                if (ExcelUtil.GetCellData(row, col + 6, sheetname) != "")
                {
                    Perform.EnterText(".//*[@id='P_L_V_v39w9_t20_c0w0_PC_t0i"+i+"_LengthInsNumeric_LengthInsNumeric']", ExcelUtil.GetCellData(row, col + 6, sheetname));
                }
                if (ExcelUtil.GetCellData(row, col + 7, sheetname) != "")
                {
                    Perform.EnterText(".//*[@id='P_L_V_v39w9_t20_c0w0_PC_t0i"+i+"_RatedSpeedInsNumeric_RatedSpeedInsNumeric']", ExcelUtil.GetCellData(row, col + 7, sheetname));
                }
                if (ExcelUtil.GetCellData(row, col + 8, sheetname) != "")
                {
                    Perform.EnterText(".//*[@id='P_L_V_v39w9_t20_c0w0_PC_t0i"+i+"_CostNewInsNumeric_CostNewInsNumeric']", ExcelUtil.GetCellData(row, col + 8, sheetname));
                }
                if (ExcelUtil.GetCellData(row, col + 9, sheetname) != "")
                {
                    Perform.EnterText(".//*[@id='P_L_V_v39w9_t20_c0w0_PC_t0i"+i+"_DescriptionInsTextBox']", ExcelUtil.GetCellData(row, col + 9, sheetname));
                }
                if (ExcelUtil.GetCellData(row, col + 10, sheetname) != "")
                {
                    Perform.EnterText(".//*[@id='P_L_V_v39w9_t20_c0w0_PC_t0i"+i+"_PremiumInsNumeric_PremiumInsNumeric']", ExcelUtil.GetCellData(row, col + 10, sheetname));
                }
                if (ExcelUtil.GetCellData(row, col + 11, sheetname) != "")
                {
                    Perform.Click(".//*[@id='P_L_V_v39w9_t20_c0w0_PC_t0i"+i+"_OtherOwnerInsCheckBox']");
                }
                //Motor
                if (ExcelUtil.GetCellData(row, col + 12, sheetname) != "")
                {
                    Perform.SelectTextDropDown(".//*[@id='P_L_V_v39w9_t20_c0w0_PC_t0i"+i+"_TypeMotorInsCombo_D_I']", ExcelUtil.GetCellData(row, col + 12, sheetname));
                }
                if (ExcelUtil.GetCellData(row, col + 13, sheetname) != "")
                {
                    Perform.EnterText(".//*[@id='P_L_V_v39w9_t20_c0w0_PC_t0i"+i+"_YearMotorInsNumeric_YearMotorInsNumeric']", ExcelUtil.GetCellData(row, col + 13, sheetname));
                }
                if (ExcelUtil.GetCellData(row, col + 14, sheetname) != "")
                {
                    Perform.EnterText(".//*[@id='P_L_V_v39w9_t20_c0w0_PC_t0i"+i+"_ManufacturerMotorInsTextBox']", ExcelUtil.GetCellData(row, col + 14, sheetname));
                }
                if (ExcelUtil.GetCellData(row, col + 15, sheetname) != "")
                {
                    Perform.EnterText(".//*[@id='P_L_V_v39w9_t20_c0w0_PC_t0i"+i+"_ModelMotorInsTextBox']", ExcelUtil.GetCellData(row, col + 15, sheetname));
                }
                if (ExcelUtil.GetCellData(row, col + 16, sheetname) != "")
                {
                    Perform.EnterText(".//*[@id='P_L_V_v39w9_t20_c0w0_PC_t0i"+i+"_SerialNumberMotorInsTextBox']", ExcelUtil.GetCellData(row, col + 16, sheetname));
                }
                if (ExcelUtil.GetCellData(row, col + 17, sheetname) != "")
                {
                    Perform.EnterText(".//*[@id='P_L_V_v39w9_t20_c0w0_PC_t0i"+i+"_CostNewMotorInsNumeric_CostNewMotorInsNumeric']", ExcelUtil.GetCellData(row, col + 17, sheetname));
                }
                //RV/Watercraft Level Coverages
                if (ExcelUtil.GetCellData(row, col + 18, sheetname) != "")
                {
                    Perform.SelectTextDropDown(".//*[@id='P_L_V_v39w9_t20_c0w0_PC_t0i"+i+"_RVWatercraftDynamicCoveragesTableLayoutPanel_12184_70097_12184_70097_Deductible1Limit_D_I']", ExcelUtil.GetCellData(row, col + 18, sheetname));
                }
                if (ExcelUtil.GetCellData(row, col + 19, sheetname) != "")
                {
                    Perform.SelectTextDropDown(".//*[@id='P_L_V_v39w9_t20_c0w0_PC_t0i"+i+"_RVWatercraftDynamicCoveragesTableLayoutPanel_12108_294_12108_294_MainLimitLimit_D_I']", ExcelUtil.GetCellData(row, col + 19, sheetname));
                }
                if (ExcelUtil.GetCellData(row, col + 20, sheetname) != "")
                {
                    Perform.Click(".//*[@id='P_L_V_v39w9_t20_c0w0_PC_t0i"+i+"_RVWatercraftDynamicCoveragesTableLayoutPanel_12147_20167_12147_20167_MainLimitLimit']");
                }
                if (ExcelUtil.GetCellData(row, col + 21, sheetname) != "")
                {
                    Perform.Click(".//*[@id='P_L_V_v39w9_t20_c0w0_PC_t0i"+i+"_RVWatercraftDynamicCoveragesTableLayoutPanel_12241_80149_12241_80149_MainLimitLimit']");
                }
                Perform.Click(".//*[@id='SaveToolStripButtonMiddle']");
            }
            col = col +24;
        }
        
    }
}
