using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.WorkflowServices;
namespace TestingApp
{
    public static class WorkflowUtility
    {
        public static void GetWorkflowReport(string listName, List oList, ClientContext clientContext)
        {
            try
            {
                //Get all list items
                CamlQuery camlQuery = new CamlQuery();
                camlQuery.ViewXml =
                    "<View><Query>" +
                        "<Where>" +
                            "<And>" +
                                "<Neq><FieldRef Name=\"Stage\"/><Value Type=\"Text\">Complete</Value></Neq>" +
                                "<Neq><FieldRef Name=\"Stage\"/><Value Type=\"Text\">Draft</Value></Neq>" +
                            "</And>" +
                        "</Where>" +
                    // "<OrderBy><FieldRef Name=\"Modified\" Ascending=\"FALSE\"/></OrderBy>" +
                    "</Query>" +
                    //"<RowLimit>5000</RowLimit>" +
                    "</View>";
                ListItemCollection collListItem = oList.GetItems(camlQuery);

                clientContext.Load(collListItem,
                     items => items.Include(
                        item => item.Id,
                        item => item["Title"],
                        item => item.ParentList));
                clientContext.ExecuteQuery();
                List<CSVOutput> list = new List<CSVOutput>();

                foreach (ListItem oListItem in collListItem)
                {
                    WorkflowInstanceCollection allinstances = WorkflowExtensions.GetWorkflowInstances(oList.ParentWeb, oListItem);
                    //clientContext.Load(allinstances);
                    //clientContext.ExecuteQuery();

                    foreach (WorkflowInstance instance in allinstances)
                    {
                        
                        if (instance.Status == WorkflowStatus.Suspended || instance.Status == WorkflowStatus.Terminated )
                        {
                            WorkflowSubscription subscription = WorkflowExtensions.GetWorkflowSubscription(oList.ParentWeb, instance.WorkflowSubscriptionId);
                            string itemUrl = String.Format(@"{0}/Lists/{1}/DispForm.aspx?ID={2}", oList.ParentWeb.Url, listName, oListItem.Id);

                            string WorkflowUrl = String.Format(@"{0}/_layouts/15/workflow.aspx?List={1}&ID={2}", oList.ParentWeb.Url, oList.Id, oListItem.Id); ;

                            Console.WriteLine(string.Format("Item ID:{0}, Title:{1}, Item URL:{2} ", oListItem.Id, oListItem["Title"], itemUrl));
                            Console.Write(string.Format("Workflow Name:{0},Workflow Status:{1}, message:{2}, Workflow URL: {3} ", subscription.Name, instance.Status, instance.FaultInfo,WorkflowUrl));
                            
                            CSVOutput item = new CSVOutput(oListItem.Id.ToString(), oListItem["Title"].ToString(), itemUrl, subscription.Name, instance.Status.ToString(), instance.FaultInfo, WorkflowUrl);
                            list.Add(item);
                        }
                    }
                }

                WriteCSV(list);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        static void WriteCSV(List<CSVOutput> list)
        {
            string filepath = ConfigurationManager.AppSettings["csvOutputPath"]==null?string.Empty:ConfigurationManager.AppSettings["csvOutputPath"];
            if (!string.IsNullOrEmpty(filepath))
            {
                var fileName = filepath + "WorkflowReport " + DateTime.Now.ToString("dd-MMM-yyyy hh-mm tt") + ".csv";
                using (CsvFileWriter writer = new CsvFileWriter(fileName))
                {
                    // Write sample data to CSV file
                    string[] str = { "ItemId", "ItemTitle", "ItemUrl", "Workflow Name",
                        "Workflow Status","Workflow Message","Workflow URL" };
                    CsvRow row = new CsvRow();
                    row.AddRange(str);
                    writer.WriteRow(row);
                    foreach (CSVOutput item in list)
                    {
                        row = new CsvRow{};
                        row.Add(item.ItemId);
                        row.Add(item.ItemTitle);
                        row.Add(item.ItemUrl);
                        row.Add(item.WorkflowName);
                        row.Add(item.WorkflowStatus);
                        row.Add(item.WorkflowMessage);
                        row.Add(item.WorkflowURL);
                        writer.WriteRow(row);
                    }
                }
            }
        }

    }
}
