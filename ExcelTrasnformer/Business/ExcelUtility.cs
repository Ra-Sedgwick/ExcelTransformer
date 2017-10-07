using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Cells;
using System.IO;
using System.Diagnostics;

namespace ExcelTrasnformer.Business
{
    public class ExcelUtility
    {

        public Workbook FileReader(String filePath)
        {
            Workbook wb = GetWorkBook(filePath);

            if (wb == null)
            {
                throw new FileNotFoundException();
            }

            Worksheet sheet = wb.Worksheets[0];

            // Consolidate Date Range
            List<int> lengthInDays = new List<int>();
            List<DateTime> statusChanged = new List<DateTime>();

            List<double> deadlineUnix = GetColumn<double>(1, 7, sheet);
            List<double> statusChangedUnix = GetColumn<double>(1, 8, sheet); 
            List<double> launchUnix = GetColumn<double>(1, 9, sheet);

            for (int i = 0; i < deadlineUnix.Count; i++)
            {
                // duration = deadline(seconds) - launch(seconds)
                // duration; seconds => days
                double lengthSeconds =  deadlineUnix[i] - launchUnix[i];
                int lengthDays = (int)lengthSeconds / 60 / 60 / 24;

                lengthInDays.Add(lengthDays);

                // Unix => Utc
                DateTime date = DateTimeOffset.FromUnixTimeSeconds((long)statusChangedUnix[i])
                    .UtcDateTime;

                statusChanged.Add(date); 
            }

            
            
       


            return wb; 
        }

        //TODO: Handle Generic Data Type. 
        private List<T> GetColumn<T>(int rowIndex, int columnIndex, Worksheet sheet)
        {
            List<T> column = new List<T>();
            Cell current;

            while (true)
            {
                current = sheet.Cells[rowIndex, columnIndex];

                if (current.Value == null)
                    break;

                column.Add((T)current.Value);
                rowIndex++;
            }

            return column;
        }
        

        private Workbook GetWorkBook(String filePath)
        {
            Workbook wb = null; 
            FileStream fileStream;

            try
            {
                fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                wb = new Workbook(fileStream);

            }
            catch (IOException e)
            {
                Debug.WriteLine(e.Message); 
            }

            return wb; 
        }
    }
}
