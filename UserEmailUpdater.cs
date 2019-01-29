using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;

namespace TestingApp
{
    public class UserEmailUpdater
    {
        public bool UpdateUserEmailId(string userLoginName, string userEmailId, ClientContext clientContext)
        {
            bool isupdated = false;
            try
            {
                Site site = clientContext.Site;
                clientContext.Load(site);
                clientContext.ExecuteQuery();
                Web web = site.RootWeb;
                clientContext.Load(web);
                clientContext.ExecuteQuery();
                var user = web.EnsureUser(userLoginName);
                user.Email = userEmailId;
                user.Update();
                clientContext.ExecuteQuery();
                user = web.EnsureUser(userLoginName);
                clientContext.Load(user);
                clientContext.ExecuteQuery();
                isupdated =  true;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return isupdated;
        }
    }
}
