using System;
using System.Collections.Generic;
using System.Linq;
using System.Security;
using System.Web;
using Microsoft.SharePoint.Client;
using System.Data;
using System.Diagnostics;

/// <summary>
/// Class that allows connection and communication with Microsoft Sharepoint server
/// 
/// 
/// Required NuGet Assemblies (DLLs in Bin folder)
/// 
/// Microsoft.SharePointOnline.CSOM version 16.1.5813.1200
/// Microsoft.IdentityModel 6.1.7600.16394
/// Microsoft.CrmSdk.Extensions 7.1.0
/// Microsoft.CrmSdk.Deployment 8.1.0.2
/// Microsoft.CrmSdk.CoreAssemblies 8.1.0.2
/// </summary>
/// 

namespace SilverWare
{
    public class SharePoint
    {
        public static SecureString getSecureString(string normalString)
        {
            //creates secure string from normal string
            SecureString secureString = new SecureString();

            foreach (char c in normalString)
            {
                secureString.AppendChar(c);
            }
            return secureString;
        }


        public static Microsoft.SharePoint.Client.ClientContext getSharepointContext(string siteUrl, string userName, string userPassword)
        {
            //returns execution context object e.g.SharepointContext

            ClientContext context = new ClientContext(siteUrl);
            context.Credentials = System.Net.CredentialCache.DefaultNetworkCredentials;
            context.Credentials = new SharePointOnlineCredentials(userName, getSecureString(userPassword));
            return context;
        }

        public static DataTable getSharepointFolderFileList(ClientContext context, Web web, string folderName)
        {
            //return filenames in folder (provided as folderName)
            DataTable dt = new DataTable();
            DataColumn col1 = new DataColumn("File");
            DataColumn col2 = new DataColumn("Folder");
            DataColumn col3 = new DataColumn("Id");
            DataColumn col4 = new DataColumn("UniqueId");

            col1.DataType = System.Type.GetType("System.String");
            col2.DataType = System.Type.GetType("System.String");
            col3.DataType = System.Type.GetType("System.String");
            col4.DataType = System.Type.GetType("System.String");

            dt.Columns.Add(col1);
            dt.Columns.Add(col2);
            dt.Columns.Add(col3);
            dt.Columns.Add(col4);

            Microsoft.SharePoint.Client.List myList = context.Web.Lists.GetByTitle(folderName);//"Temp"

            CamlQuery camlQuery = new CamlQuery();
            camlQuery.ViewXml = @"<View Scope = 'Recursive'>
                                    <Query>
                                    </Query>
                                </View>";
            ListItemCollection listItems = myList.GetItems(camlQuery);
            context.Load(listItems);
            context.ExecuteQuery();

            for (int i = 0; i < listItems.Count; i++)
            {
                ListItem itemOfIntererest = listItems[i];
                DataRow row = dt.NewRow();
                row[col1] = itemOfIntererest.FieldValues["FileRef"];
                row[col2] = itemOfIntererest.FieldValues["FileDirRef"];
                row[col3] = itemOfIntererest.FieldValues["ID"];
                row[col4] = itemOfIntererest.FieldValues["UniqueId"];
                dt.Rows.Add(row);
            }
            return dt;
        }


        public static DataTable getSharepointWebLists(ClientContext context, Web web)
        {
            //returns Sharepoint lists

            context.Load(web.Lists, lists => lists.Include(list => list.EntityTypeName, list => list.BaseType, list => list.Title, list => list.Id, list => list.DefaultDisplayFormUrl, list => list.ItemCount)); // For each list, retrieve Title and Id. 
            context.ExecuteQuery();

            var listItemCount = web.Lists.Count;

            DataTable table = new DataTable();
            DataColumn col1 = new DataColumn("Folder name");
            DataColumn col2 = new DataColumn("File count");
            DataColumn col3 = new DataColumn("Id");

            col1.DataType = System.Type.GetType("System.String");
            col2.DataType = System.Type.GetType("System.String");
            col3.DataType = System.Type.GetType("System.String");

            table.Columns.Add(col1);
            table.Columns.Add(col2);
            table.Columns.Add(col3);

            foreach (Microsoft.SharePoint.Client.List list in web.Lists)
            {

                //Response.Write(list.EntityTypeName + "    " + list.BaseType + "     " + list.Title + "               " + "Id:" + list.Id + "   Item count:" + list.ItemCount + "  " + list.DefaultDisplayFormUrl + "<br/>");
                DataRow row = table.NewRow();
                row[col1] = list.Title;
                row[col2] = list.ItemCount;
                row[col3] = list.Id;
                table.Rows.Add(row);
            }
            return table;
        }
    }

}