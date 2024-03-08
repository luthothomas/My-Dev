
using ToolTest.DataFile;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.IO;
using System.Configuration;
using System.Net;

namespace ToolTest
{
    class Program
    {
        static string urlAddress = "https://clientportal.jse.co.za/_layouts/15/DownloadHandler.ashx?FileName=/YieldX/Derivatives/Docs_DMTM";

        static void Main(string[] args)
        {
            ToolTest.DataFile.FileClass.CrearteFile(@"C:\Temp\Report2023\");
            //build file name by creating the date to get the file name as expected.
            DateTime today = new DateTime();
            DateTime answer = today.AddDays(1);
            for (int i = 0; i < 365; i++)
            {
                //Set the start date.
                today = new DateTime(2023, 01, 01, 0, 00, 00);
                answer = today.AddDays(i);
                string date1 = answer.ToString("yyyyMMdd");
                // get the file name with the prix _D_Daily MTM Report.xls
                string fileName = date1 + "_D_Daily MTM Report.xls";

                if (!File.Exists(@"C:\Temp\Report2023\" + fileName))
                {
                    //Download file to the local path.
                    ToolTest.DataFile.FileClass.DownLoadFile(@urlAddress + "/" + fileName, @"C:\Temp\Report2023\" + fileName);
                   
                    //Transform file into datatable and return.
                    DataTable dt = ToolTest.DataFile.FileClass.ReadExcel(fileName);

                    //Load data file into the database.
                    ToolTest.DataFile.FileClass.BulkInsert(dt);
                    //List the successful file dwnloaded and saved into db.
                    Console.WriteLine(@"C:\Temp\Report2023\" + fileName);
                }
                else
                {
                    Console.WriteLine("File already exist");
                }
            }
            Console.WriteLine("Download was successful. Please enter to exist.");
            Console.ReadKey();
        }
    }
}
