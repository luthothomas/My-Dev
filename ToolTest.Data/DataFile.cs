using System;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Net;
using Excel = Microsoft.Office.Interop.Excel;

namespace ToolTest.DataFile
{
    public class FileClass
    {
        private static bool ValidateDate(string strDate)
        {
            DateTime value;
            if (!DateTime.TryParse(strDate, out value))
            {
                return false;
            }
            return true;
        }


        public static DataTable ReadExcel(string filepath)
        {
            DataTable dt = new DataTable();
            dt.Clear();
            dt.Columns.Add("Contract");
            dt.Columns.Add("ExpiryDate");
            dt.Columns.Add("Classification");
            dt.Columns.Add("MTMYield");
            dt.Columns.Add("MarkPrice");
            dt.Columns.Add("SpotRate");
            dt.Columns.Add("PreviousMTM");
            dt.Columns.Add("PreviousPrice");
            dt.Columns.Add("PremiumOnOption");
            dt.Columns.Add("Volatility");
            dt.Columns.Add("Delta");
            dt.Columns.Add("DeltaValue");
            dt.Columns.Add("ContractsTraded");
            dt.Columns.Add("OpenInterest");

            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkBook = xlApp.Workbooks.Open(@"C:\Temp\Report2023\" + filepath);
            Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            Excel.Range xlRange = xlWorkSheet.UsedRange;
            int totalRows = xlRange.Rows.Count;
            int totalColumns = xlRange.Columns.Count;

            DataRow row = null;
            ToolTest.DTO.DailMReport report = new DTO.DailMReport();

            for (int rowCount = 6; rowCount <= totalRows; rowCount++)
            {

                report.Contract = Convert.ToString((xlRange.Cells[rowCount, 1] as Excel.Range).Text);
                report.ExpiryDate = Convert.ToString((xlRange.Cells[rowCount, 3] as Excel.Range).Text);
                report.Classification = Convert.ToString((xlRange.Cells[rowCount, 4] as Excel.Range).Text);
                report.MTMYield = Convert.ToString((xlRange.Cells[rowCount, 5] as Excel.Range).Text);
                report.MarkPrice = Convert.ToString((xlRange.Cells[rowCount, 6] as Excel.Range).Text);
                report.SpotRate = Convert.ToString((xlRange.Cells[rowCount, 7] as Excel.Range).Text);
                report.PreviousMTM = Convert.ToString((xlRange.Cells[rowCount, 8] as Excel.Range).Text);
                report.PreviousPrice = Convert.ToString((xlRange.Cells[rowCount, 9] as Excel.Range).Text);
                report.PremiumOnOption = Convert.ToString((xlRange.Cells[rowCount, 10] as Excel.Range).Text);
                report.Volatility = Convert.ToString((xlRange.Cells[rowCount, 11] as Excel.Range).Text);
                report.Delta = Convert.ToString((xlRange.Cells[rowCount, 12] as Excel.Range).Text);
                report.DeltaValue = Convert.ToString((xlRange.Cells[rowCount, 13] as Excel.Range).Text);
                report.ContractsTraded = Convert.ToString((xlRange.Cells[rowCount, 14] as Excel.Range).Text);
                report.OpenInterest = Convert.ToString((xlRange.Cells[rowCount, 15] as Excel.Range).Text);

                if (report.Contract != null && report.Contract != string.Empty)
                {
                    row = dt.NewRow();
                    row["Contract"] = report.Contract;
                    row["ExpiryDate"] = report.ExpiryDate;
                    row["Classification"] = report.Classification;
                    row["MTMYield"] = report.MTMYield;
                    row["MarkPrice"] = report.MarkPrice;
                    row["SpotRate"] = report.SpotRate;
                    row["PreviousMTM"] = report.PreviousMTM;
                    row["PreviousPrice"] = report.PreviousPrice;
                    row["PremiumOnOption"] = report.PremiumOnOption;
                    row["Volatility"] = report.Volatility;
                    row["Delta"] = report.Delta;
                    row["DeltaValue"] = report.DeltaValue;
                    row["ContractsTraded"] = report.ContractsTraded;
                    row["OpenInterest"] = report.OpenInterest;

                    dt.Rows.Add(row);

                }
            }
            xlWorkBook.Close();
            xlApp.Quit();
            // Console.Read();
            GC.Collect();

            return dt;
        }


        public static void BulkInsert(DataTable dt)
        {
           
            using (var con = new SqlConnection(ConfigurationManager.ConnectionStrings["DBConnection"].ConnectionString))
            {
            
                con.Open();

                // Get a reference to a single row in the table. 
                DataRow[] rowArray = dt.Select();

                using (SqlBulkCopy bulkCopy = new SqlBulkCopy(con))
                {
                    bulkCopy.DestinationTableName = "dbo.DailyMTM";

                    try
                    {
                        // Write the array of rows to the destination.
                        bulkCopy.WriteToServer(rowArray);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }

                }

            }//using
        }


        public static void CrearteFile(string path)
        {

            try
            {
                bool exists = System.IO.Directory.Exists(path);

                if (!exists)
                    System.IO.Directory.CreateDirectory(path);

            }
            catch (Exception e)
            {
                Console.WriteLine("The process failed: {0}", e.ToString());
            }
        }

        public static void DownLoadFile(string url, string urlStringFile)
        {
            using (WebClient client = new WebClient())
            {
                client.DownloadFile(@url, @urlStringFile);

                String ver = File.ReadAllText(urlStringFile);
                string test = ver.ToString();
            }
        }

    }
}
 

