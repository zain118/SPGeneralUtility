using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Sharing;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;
using TestingApp.CSV;

namespace TestingApp
{
    public class InviteExternalUser
    {
        public static void InviteExtUser()
        {
            List<CSVOutputExternalUser> externalUserColl = ReadCSV();
            Dictionary<string, Role> roleDictionary = new Dictionary<string, Role>()
            {
                {"None", Role.None },
                {"View", Role.View },
                {"Edit", Role.Edit },
                {"Owner", Role.Owner },
                {"LimitedView", Role.LimitedView },
                {"LimitedEdit", Role.LimitedEdit },
                {"Review", Role.Review },
                {"RestrictedView", Role.RestrictedView }
            };

            string username = ConfigurationManager.AppSettings["username"];
            string password = ConfigurationManager.AppSettings["password"];
            var siteDomain = ConfigurationManager.AppSettings["siteDomain"];
            string emailMessage;
            using (var securePassword = new SecureString())
            {
                foreach (var c in password.ToCharArray())
                    securePassword.AppendChar(c);
                var count = 0;
                foreach (var extUser in externalUserColl.Skip(1))
                {

                    emailMessage = "Dear " + extUser.Name + Environment.NewLine +
                        ", Welcome to SharePoint site, please follow the below link to request access.";

                    using (var ctx = new ClientContext(siteDomain + extUser.Site))
                    {
                        try
                        {
                            count++;
                            ctx.Credentials = new SharePointOnlineCredentials(username, securePassword);
                            var web = ctx.Web;
                            ctx.Load(web, x => x.SiteGroups);
                            ctx.ExecuteQuery();

                            var users = new List<UserRoleAssignment>();
                            //Role role = roleDictionary[extUser.PermissionGroup];
                            users.Add(new UserRoleAssignment()
                            {
                                UserId = extUser.Email,
                                Role = Role.LimitedView

                            });
                            WebSharingManager.UpdateWebSharingInformation(ctx, web, users, true, emailMessage, true, true);
                            ctx.ExecuteQuery();
                            Console.WriteLine("Completed " + count);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message + count);
                        }
                    }
                }
                Console.ReadLine();
            }
        }
        static List<CSVOutputExternalUser> ReadCSV()
        {
            // Read sample data from CSV file
            List<CSVOutputExternalUser> externalUserCollection = new List<CSVOutputExternalUser>();
            using (CSVFileReader reader = new CSVFileReader(ConfigurationManager.AppSettings["csvInputPath"]))
            {
                CsvRow row = new CsvRow();
                while (reader.ReadRow(row))
                {
                    CSVOutputExternalUser ext = new CSVOutputExternalUser(row[1], row[0], row[3], row[2]);
                    externalUserCollection.Add(ext);
                }
            }
            Console.WriteLine(externalUserCollection.Count);
            return externalUserCollection;
        }
    }
}
