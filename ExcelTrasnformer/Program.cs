using Aspose.Cells;
using ExcelTrasnformer.Business;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelTrasnformer
{
    class Program
    {
        static void Main(string[] args)
        {
            ExcelUtility  excel = new ExcelUtility();
            Workbook workBook = excel.TransformWorkbook("C:/Data/train.xlsx");

            Console.ReadKey();
        }
    }
}
