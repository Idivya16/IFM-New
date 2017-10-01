using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DIA_PPA_NB
{
    class PageUtility
    {
        public static void Calendar(int row,int col,string sheetname)
        { 
        String month = ExcelUtil.GetCellData(row, col, sheetname);
        int month1 = Int32.Parse(month) - 1;
        string year = ExcelUtil.GetCellData(row, col+2, sheetname);
        string yr = year.Substring(3, 1);
        int yr1 = Int32.Parse(yr);
        Console.WriteLine(yr);
            string date = ExcelUtil.GetCellData(row, col+1, sheetname);
        //Enter month
        Perform.Click(".//*[@id='P_L_ClientSubmissionWithAddressAndPOBox_Client1InsName_BirthDate_BirthDate_DDD_C_FNP_M" + month1 + "']");
            Console.WriteLine("Month is selected");
            //Enter year
            string currentyear = DateTime.Now.Year.ToString();
        string diffyear = (Int32.Parse(currentyear) - Int32.Parse(year)).ToString();
        string diff = diffyear.Substring(0, 1);
            for (int i = 0; i<Int32.Parse(diff); i++)
            {
                Perform.driver.FindElement(By.CssSelector("img.dxEditors_edtCalendarFNPrevYear_Metropolis")).Click();
    }

    Perform.Click(".//*[@id='P_L_ClientSubmissionWithAddressAndPOBox_Client1InsName_BirthDate_BirthDate_DDD_C_FNP_Y" + yr1 + "']");
            Console.WriteLine("Year is selected" + year);
            Perform.Click(".//*[@id='P_L_ClientSubmissionWithAddressAndPOBox_Client1InsName_BirthDate_BirthDate_DDD_C_FNP_BO']");
            IList<IWebElement> allDates = Perform.driver.FindElements(By.XPath("//td[contains(@class,'dxeCalendarDay_Metropolis')]"));
            foreach (IWebElement ele in allDates)
            {

                String date1 = ele.Text;

                if (date1.Equals(date))
                {
                    ele.Click();
                    break;
                }

            }

        }

        public static void Calendardriver(int row,int col,int j, string sheetname)
        {
            
            String month = ExcelUtil.GetCellData(row, col, sheetname);
            int month1 = Int32.Parse(month) - 1;
            string year = ExcelUtil.GetCellData(row, col+2, sheetname);
            string yr = year.Substring(3, 1);
            int yr1 = Int32.Parse(yr);
            Console.WriteLine(yr);
            string date = ExcelUtil.GetCellData(row, col+1, sheetname);
            //Enter month
            Perform.Click(".//*[@id='P_L_V_v33w9_t15_c0w0_PC_t1i"+j+"_c0w0_t0_InsName_BirthDate_BirthDate_DDD_C_FNP_M" +month1+ "']");
            Console.WriteLine("Month is selected");
            //Enter year
            string currentyear = DateTime.Now.Year.ToString();
            string diffyear = (Int32.Parse(currentyear) - Int32.Parse(year)).ToString();
            string diff = diffyear.Substring(0, 1);
            for (int i = 0; i < Int32.Parse(diff); i++)
            {
                Perform.driver.FindElement(By.CssSelector("img.dxEditors_edtCalendarFNPrevYear_Metropolis")).Click();
                if (diff=="9")
            {
                    Perform.driver.FindElement(By.CssSelector("img.dxEditors_edtCalendarFNPrevYear_Metropolis")).Click();
                }
                System.Threading.Thread.Sleep(200);
            }

            Perform.Click(".//*[@id='P_L_V_v33w9_t15_c0w0_PC_t1i"+j+"_c0w0_t0_InsName_BirthDate_BirthDate_DDD_C_FNP_Y" +yr1+ "']");
            Console.WriteLine("Year is selected" + year);
            Perform.Click(".//*[@id='P_L_V_v33w9_t15_c0w0_PC_t1i"+j+"_c0w0_t0_InsName_BirthDate_BirthDate_DDD_C_FNP_BO']");
            IList<IWebElement> allDates = Perform.driver.FindElements(By.XPath("//td[contains(@class,'dxeCalendarDay_Metropolis')]"));
            foreach (IWebElement ele in allDates)
            {

                String date1 = ele.Text;

                if (date1.Equals(date))
                {
                    ele.Click();
                    break;
                }

            }

        }
        public static void policyrelation(int row,int i,int col,string sheetname)
        {
            
            string relpol = ExcelUtil.GetCellData(row,col+17, sheetname);
            if(relpol=="Policyholder")
            {
                System.Threading.Thread.Sleep(500);
                Perform.mouseclick(".//*[@id='P_L_V_v33w9_t15_c0w0_PC_t1i"+i+"_c0w0_t0_RelationToPolicyHolderInsCombo_D_DDD_L_LBI9T0']");
                
            }
            if (relpol == "Spouse of Policyholder")
            {
                System.Threading.Thread.Sleep(500);
                Perform.mouseclick(".//*[@id='P_L_V_v33w9_t15_c0w0_PC_t1i"+i+"_c0w0_t0_RelationToPolicyHolderInsCombo_D_DDD_L_LBI10T0']");
            }
            if (relpol == "Child of Policyholder")
            {
                System.Threading.Thread.Sleep(500);
                Perform.mouseclick(".//*[@id='P_L_V_v33w9_t15_c0w0_PC_t1i"+i+"_c0w0_t0_RelationToPolicyHolderInsCombo_D_DDD_L_LBI3T0']");
            }
            if (relpol == "Policyholder #2")
            {
                System.Threading.Thread.Sleep(500);
                Perform.mouseclick(".//*[@id='P_L_V_v33w9_t15_c0w0_PC_t1i"+i+"_c0w0_t0_RelationToPolicyHolderInsCombo_D_DDD_L_LBI6T0']");
            }
        }



    }
}
