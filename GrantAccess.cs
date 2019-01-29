using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using Microsoft.SharePoint.Client;
using System.Net.Mail;
using Microsoft.SharePoint.Client.Utilities;
using System.Security;
using TestingApp.CSV;

namespace TestingApp
{
    public class GrantAccess
    {
        public static List<CSVOutputExternalUser> ReadCSV()
        {
            // Read sample data from CSV file
            List<CSVOutputExternalUser> externalUserCollection = new List<CSVOutputExternalUser>();
            using (CSVFileReader reader = new CSVFileReader(ConfigurationManager.AppSettings["csvInputPath"]))
            {
                CsvRow row = new CsvRow();
                while (reader.ReadRow(row))
                {
                    CSVOutputExternalUser ext = new CSVOutputExternalUser(row[1],row[0],row[3],row[2]);
                    externalUserCollection.Add(ext);
                }
            }
            Console.WriteLine(externalUserCollection.Count);
            return externalUserCollection;
        }

        public static void WriteCSV(List<CSVOutputExternalUser> list)
        {
            string filepath = ConfigurationManager.AppSettings["csvOutputPath"] == null ? string.Empty : ConfigurationManager.AppSettings["csvOutputPath"];
            if (!string.IsNullOrEmpty(filepath))
            {
                var fileName = filepath + "GrantAccessReport " + DateTime.Now.ToString("dd-MMM-yyyy hh-mm tt") + ".csv";
                using (CsvFileWriter writer = new CsvFileWriter(fileName))
                {
                    // Write sample data to CSV file
                    string[] str = { "Name", "Email", "Site", "Permission Group", "isAdded" };
                    CsvRow row = new CsvRow();
                    row.AddRange(str);
                    writer.WriteRow(row);
                    foreach (CSVOutputExternalUser item in list)
                    {
                        row = new CsvRow();
                        row.Add(item.Name);
                        row.Add(item.Email);
                        row.Add(item.Site);
                        row.Add(item.PermissionGroup);
                        row.Add(item.Message);
                        writer.WriteRow(row);
                    }
                }
            }
        }

        public static void SendMail(string email, string name, string site, string username, SecureString password)
        {
            string emailMessage;
            var siteDomain = ConfigurationManager.AppSettings["siteDomain"];
            string targetSite = siteDomain + site;



            emailMessage = "<table border='0' cellspacing='0' cellpadding='0'><tbody>" +
                            "<tr><td style= 'padding:12px 0 14px 0;font-family:Segoe UI,Arial,sans-serif;font-size:16px;color:#3D3D3D' ><span style='font-size:16px;'>Dear " + name + ",</td></tr>" +
                            "<tr><td style= 'padding:12px 0 18px 0;font-family:Segoe UI,Arial,sans-serif;font-size:16px;color:#3D3D3D'><span style='font-size:16px;' > Welcome to the Sharepoint site, your access has now been granted.</span></td></tr>" +
                            "<tr><td style = 'padding:6px 0 48px 0;font-family:Segoe UI,Arial,sans-serif';color:'#3D3D3D'><span style='font-size:32px;'>Go To <a href = \"" + targetSite + "\" target =\"_blank\"><span style='color:darkorange;text-decoration-color:darkorange;text-decoration:underline; '> SharePoint Site </span></a></span></td></tr>" +
                            "</tbody></table>";

            

            var emailp = new EmailProperties();
            emailp.To = new List<string> { email };
            emailp.CC = new List<string> { "abcd@xyz.com" };
            emailp.From = "no-reply@sharepointonline.com";
            emailp.Body = emailMessage;
            emailp.Subject = "Access Granted";
            using (var ctx = new ClientContext(siteDomain + site))
            {
                ctx.Credentials = new SharePointOnlineCredentials(username, password);
                var web = ctx.Web;
                Utility.SendEmail(ctx, emailp);
                ctx.ExecuteQuery();
            }
        }
        public static void GrantAccessToExt()
        {
            List<CSVOutputExternalUser> externalUserColl = ReadCSV();

            string username = ConfigurationManager.AppSettings["username"];
            string password = ConfigurationManager.AppSettings["password"];
            var siteDomain = ConfigurationManager.AppSettings["siteDomain"];

            using (var securePassword = new SecureString())
            {
                foreach (var c in password.ToCharArray())
                    securePassword.AppendChar(c);

                List<CSVOutputExternalUser> list = new List<CSVOutputExternalUser>();
                var count = 0;
                foreach (var extUser in externalUserColl.Skip(1))
                {
                    using (var ctx = new ClientContext(siteDomain + extUser.Site))
                    {
                        ctx.Credentials = new SharePointOnlineCredentials(username, securePassword);
                        var web = ctx.Web;
                        GroupCollection groups = web.SiteGroups;
                        //web.Lists.
                        string perm = extUser.PermissionGroup;
                        ctx.Load(groups, groupitems => groupitems.Include(groupitem => groupitem.Title).Where(groupitem => groupitem.Title == perm));

                        try
                        {
                            count++;
                            var user = ctx.Web.EnsureUser(extUser.Email);
                            ctx.Load(user);
                            ctx.ExecuteQuery();
                            var spUserToAdd = groups[0].Users.AddUser(user);
                            ctx.Load(spUserToAdd);
                            ctx.ExecuteQuery();

                            CSVOutputExternalUser item = new CSVOutputExternalUser(extUser.Email, extUser.Name, extUser.Site, perm, "Added");
                            list.Add(item);

                            SendMail(extUser.Email, extUser.Name, extUser.Site, username, securePassword);
                            Console.WriteLine("Completed " + count);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message + " " + count);
                            if (ex.Message == "Failure sending mail.")
                            {
                                CSVOutputExternalUser item = new CSVOutputExternalUser(extUser.Email, extUser.Name, extUser.Site, perm, "Added but Failure sending mail");
                                list.Add(item);
                            }
                            else
                            {
                                CSVOutputExternalUser item = new CSVOutputExternalUser(extUser.Email, extUser.Name, extUser.Site, perm, ex.Message);
                                list.Add(item);
                            }
                        }
                    }
                }
                WriteCSV(list);
            }
        }
    }
}
