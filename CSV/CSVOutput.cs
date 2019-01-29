using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestingApp
{
    public class CSVOutput
    {
        public string ItemId { get; set; }
        public string ItemTitle { get; set; }
        public string ItemUrl { get; set; }
        public string WorkflowName { get; set; }
        public string WorkflowStatus { get; set; }
        public string WorkflowMessage { get; set; }
        public string WorkflowURL { get; set; }

        public CSVOutput(string ItemId, string ItemTitle, string ItemUrl, 
            string WorkflowName, string WorkflowStatus, string WorkflowMessage, string WorkflowURL)
        {
            this.ItemId = ItemId;
            this.ItemTitle = ItemTitle;
            this.ItemUrl = ItemUrl;
            this.WorkflowName = WorkflowName;
            this.WorkflowStatus = WorkflowStatus;
            this.WorkflowMessage = WorkflowMessage;
            this.WorkflowURL = WorkflowURL;
        }
    }
   
}
