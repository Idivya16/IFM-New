using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace VR_COM_BOP

{
    class ExcelUtil
    {
        private static XSSFWorkbook hssfwb;


        public static void setExcelFile(String Path)
        {
            using (FileStream file = new FileStream(Path, FileMode.Open, FileAccess.Read))
            {
                hssfwb = new XSSFWorkbook(file);
            }
        }

        public static String GetCellData(int RowNum, int ColNum, String SheetName)
        {


            ISheet sheet = hssfwb.GetSheet(SheetName);
            try
            {
                if (sheet.GetRow(RowNum).GetCell(ColNum).CellType == CellType.Numeric)
                {
                    double CellData = sheet.GetRow(RowNum).GetCell(ColNum).NumericCellValue;
                    return CellData.ToString();
                }
                if (sheet.GetRow(RowNum).GetCell(ColNum).CellType == CellType.String)
                {
                    String CellData = sheet.GetRow(RowNum).GetCell(ColNum).StringCellValue;
                    return CellData;
                }

            }
            catch (Exception e)
            {

                Console.WriteLine(e);
            }
            return "";
        }

        public static int getRowCount(String SheetName)
        {
            ISheet sheet = hssfwb.GetSheet(SheetName);
            int number = sheet.LastRowNum + 1;
            return number;
        }
    }
}


    
    

