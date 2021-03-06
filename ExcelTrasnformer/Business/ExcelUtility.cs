﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Cells;
using System.IO;
using System.Diagnostics;
using System.Runtime.Serialization.Json;
using Newtonsoft.Json;
using System.Net;
using Newtonsoft.Json.Linq;

namespace ExcelTrasnformer.Business
{
    public class ExcelUtility
    {
    public Workbook TransformWorkbook(String filePath)
        {
            #region File Work

            Workbook wb = GetWorkBook(filePath);

            if (wb == null)
                throw new FileNotFoundException();

            Worksheet sheet = wb.Worksheets[0];

            #endregion

            #region Date Consoliation
            // Consolidate Date Range
            List<double> lengthInDays = new List<double>();
            List<DateTime> statusChanged = new List<DateTime>();
            List<DateTime> deadline = new List<DateTime>();
            List<DateTime> launch = new List<DateTime>();

            List<double> deadlineUnix = GetColumn<double>(1, 7, sheet);
            List<double> statusChangedUnix = GetColumn<double>(1, 8, sheet);
            List<double> launchUnix = GetColumn<double>(1, 9, sheet);
            DateTime date;
            DateTime deadlineDate;
            DateTime launchDate;

            for (int i = 0; i < deadlineUnix.Count; i++)
            {
                // duration = deadline(seconds) - launch(seconds)
                // duration; seconds => days
                //double lengthSeconds = deadlineUnix[i] - launchUnix[i];
                //int lengthDays = (int)lengthSeconds / 60 / 60 / 24;

                //lengthInDays.Add(lengthDays);

                // Unix => Utc
                date = DateTimeOffset.FromUnixTimeSeconds((long)statusChangedUnix[i])
                    .UtcDateTime;

                statusChanged.Add(date);

                deadlineDate = DateTimeOffset.FromUnixTimeSeconds((long)deadlineUnix[i])
                                    .UtcDateTime;

                deadline.Add(deadlineDate);

                launchDate = DateTimeOffset.FromUnixTimeSeconds((long)launchUnix[i])
                    .UtcDateTime;   

                launch.Add(launchDate);

                lengthInDays.Add((deadlineDate - launchDate).TotalDays); 

            }


            #endregion

            #region Currency Consolidation

            List<String> currencys = GetColumn<String>(1, 6, sheet);
            List<int> goalAmounts = GetColumn<int>(1, 3, sheet);
            List<decimal> normalizedAmounts = new List<decimal>();
            JObject rates = GetExchangeRates();
            decimal newAmount;
            decimal rate;

            for (int i = 0; i < currencys.Count; i++)
            {
                newAmount = goalAmounts[i]; 
                if (!currencys[i].Equals("USD"))
                {
                    rate = (decimal)rates["rates"][currencys[i]];
                    newAmount *= rate;
                }

                currencys[i] = "USD"; 
                normalizedAmounts.Add(newAmount); 

            }


            #endregion
            sheet.Cells[0, 13].Value = "Length";
            for (int i = 1; i < lengthInDays.Count; i++)
            {
                sheet.Cells[i, 13].Value = lengthInDays[i - 1]; 
                
                if (!sheet.Cells[i, 6].Equals("USD"))
                {
                    sheet.Cells[i, 6].Value = "USD";
                    sheet.Cells[i, 3].Value = normalizedAmounts[i - 1]; 
                }

                sheet.Cells[i, 8].Value = statusChanged[i - 1];
                sheet.Cells[i, 7].Value = deadline[i - 1];
                sheet.Cells[i, 9].Value = launch[i - 1]; 

            }


            String newFileName = filePath.Split('.')[0] + "-formatted-Alt.xlsx";
            wb.Save(newFileName);

            return wb; 
        }

        public JObject GetExchangeRates()
        {
            var url = "http://api.fixer.io/2017-01-01?base=USD";
            var wc = new WebClient { Proxy = null };
            //var jsonString = wc.DownloadString(url);
            var jsonString = "{\"base\":\"USD\",\"date\":\"2016-12-30\",\"rates\":{\"AUD\":1.3847,\"BGN\":1.8554,\"BRL\":3.2544,\"CAD\":1.346,\"CHF\":1.0188,\"CNY\":6.9445,\"CZK\":25.634,\"DKK\":7.0528,\"GBP\":0.81224,\"HKD\":7.7555,\"HRK\":7.1717,\"HUF\":293.93,\"IDR\":13446.0,\"ILS\":3.84,\"INR\":67.92,\"JPY\":117.07,\"KRW\":1204.3,\"MXN\":20.655,\"MYR\":4.486,\"NOK\":8.62,\"NZD\":1.438,\"PHP\":49.585,\"PLN\":4.1839,\"RON\":4.306,\"RUB\":61.0,\"SEK\":9.0622,\"SGD\":1.4452,\"THB\":35.79,\"TRY\":3.5169,\"ZAR\":13.715,\"EUR\":0.94868}}";
            return JObject.Parse(jsonString);
        }

        // Read in column of spreadsheet to list. 
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
