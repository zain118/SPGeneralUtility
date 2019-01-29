using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestingApp.CSV
{
    public class CSVOutputExternalUser
    {
        public string Email { get; set; }
        public string Name { get; set; }
        public string Site { get; set; }
        public string PermissionGroup { get; set; }
        public string Message { get; set; }

        public CSVOutputExternalUser(string email, string name, string site, string permissionGroup)
        {
            this.Email = email;
            this.Name = name;
            this.Site = site;
            this.PermissionGroup = permissionGroup;
        }
        public CSVOutputExternalUser(string email, string name, string site, string permissionGroup,string message): this( email,  name,  site,  permissionGroup)
        { 
            this.Message = message;
        }
    }
}
