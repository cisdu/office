using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using System.IO;
using NPOI.SS.Util;

namespace npoi
{
    class Program
    {
        static void Main(string[] args)
        {

            IWorkbook workbook = null;  //新建IWorkbook对象  
            string fileName = "E:\\Excel2003.xlsx";
            FileStream fileStream = new FileStream(fileName, FileMode.Open, FileAccess.Read);
            if (fileName.IndexOf(".xlsx") > 0) // 2007版本  
            {
                workbook = new XSSFWorkbook(fileStream);  //xlsx数据读入workbook  
            }
            else if (fileName.IndexOf(".xls") > 0) // 2003版本  
            {
                workbook = new HSSFWorkbook(fileStream);  //xls数据读入workbook  
            }
            ISheet sheet = workbook.GetSheetAt(0);  //获取第一个工作表  
            IRow row;// = sheet.GetRow(0);            //新建当前工作表行数据  
            for (int i = 0; i < sheet.LastRowNum; i++)  //对工作表每一行  
            {
                row = sheet.GetRow(i);   //row读入第i行数据  
                if (row != null)
                {
                    for (int j = 0; j < row.LastCellNum; j++)  //对工作表每一列  
                    {
                        string cellValue = row.GetCell(j).ToString(); //获取i行j列数据  
                        Console.WriteLine(cellValue);
                    }
                }
            }

           var result= sheet.AddMergedRegion(new CellRangeAddress(0, 2, 0, 0));
            FileStream file2003 = new FileStream(@"E:\\Excel2003" +DateTime.Now.ToString("hhmmss")+".xlsx", FileMode.Create);
            workbook.Write(file2003);
            Console.ReadLine();
            fileStream.Close();
            file2003.Close();
            workbook.Close();
        }
    }
}
