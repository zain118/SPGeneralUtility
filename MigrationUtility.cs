using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestingApp
{
    class MigrationUtility
    {
        public static void CopyDocuments(string srcUrl, string srcLibrary, string destUrl, string destLibrary)
        {
            // set up the src client
            ClientContext srcContext = new ClientContext(srcUrl);
            
            // set up the destination context
            ClientContext destContext = new ClientContext(destUrl);

            // get the source list and items
            Web srcWeb = srcContext.Web;
            List srcList = srcWeb.Lists.GetByTitle(srcLibrary);
            ListItemCollection itemColl = srcList.GetItems(new CamlQuery());
            srcContext.Load(itemColl);
            srcContext.ExecuteQuery();

            // get the destination list
            Web destWeb = destContext.Web;
            destContext.Load(destWeb);
            destContext.ExecuteQuery();

            foreach (var doc in itemColl)
            {
                try
                {
                    //if (doc.FileSystemObjectType == FileSystemObjectType.File) //Field or Property "FileAttachement not found."
                    //{
                    // get the file
                    Microsoft.SharePoint.Client.File file = doc.File;
                    srcContext.Load(file);
                    srcContext.ExecuteQuery();

                    // build destination url
                    string nLocation = destWeb.ServerRelativeUrl.TrimEnd('/') + "/" + destLibrary.Replace(" ", "") + "/" + file.Name;

                    // read the file, copy the content to new file at new location
                    FileInformation fileInfo = Microsoft.SharePoint.Client.File.OpenBinaryDirect(srcContext, file.ServerRelativeUrl);
                    Microsoft.SharePoint.Client.File.SaveBinaryDirect(destContext, nLocation, fileInfo.Stream, true);
                    // }
                }
                catch (Exception ex)
                {
                    throw ex;
                }
            }
            Console.WriteLine("success...");
        }
    }
}
