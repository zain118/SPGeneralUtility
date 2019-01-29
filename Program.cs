using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;

namespace TestingApp
{
    class Program
    {
        static void Main(string[] args)
        {
            string username = ConfigurationManager.AppSettings["username"].ToString();
            string password = ConfigurationManager.AppSettings["password"].ToString();
            string siteUrl = ConfigurationManager.AppSettings["siteurl"].ToString();


            ClientContext clientContext = new ClientContext(siteUrl);
            SecureString securePassWord = new SecureString();
            foreach (var c in password) securePassWord.AppendChar(c);
            clientContext.Credentials = new SharePointOnlineCredentials(username, securePassWord);

           // InviteExternalUser.InviteExtUser();
           // GrantAccess.GrantAccessToExt();
        }
       
    }
}
