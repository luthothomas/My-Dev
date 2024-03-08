using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ToolTest.DTO
{      
    public class DailMReport
    {
        public string Contract { get; set; }
        public string ExpiryDate { get; set; }
        public string Classification { get; set; }
        public string urlAddress { get; set; } //"https://clientportal.jse.co.za/_layouts/15/DownloadHandler.ashx?FileName=/YieldX/Derivatives/Docs_DMTM";
        public string MTMYield { get; set; }
        public string MarkPrice { get; set; }
        public string SpotRate { get; set; }
        public string PreviousMTM { get; set; }
        public string PreviousPrice { get; set; }
        public string PremiumOnOption { get; set; }
        public string Volatility { get; set; }
        public string Delta { get; set; }
        public string DeltaValue { get; set; }
        public string ContractsTraded { get; set; }
        public string OpenInterest { get; set; }
    }
    
}
