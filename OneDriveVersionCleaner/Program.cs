using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Security;
using Microsoft.SharePoint.Client;

namespace OneDriveVersionCleaner
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            // ------------------------------------------------------------------------------------------------------------
            // IMPORTANT: This is not production code, and it deletes files from your cloud storage. Use at your own risk!
            // ------------------------------------------------------------------------------------------------------------

            // TODO: Configure these values.
            string webFullUrl = "https://[custom].sharepoint.com/personal/[custom]";
            string username = "[custom your identityhere]";
            SecureString password = new NetworkCredential("", "[custom]").SecurePassword;  // not secure

            string[] folderRootRelativeUrls = {
                "/Documents/Code",
				"/Documents/YourCustomFoldersHere"
            };

            bool recurse = true;

            // --- End of configuration ---

            ClientContext context = new ClientContext(webFullUrl)
            {
                Credentials = new SharePointOnlineCredentials(username, password)
            };
            context.Load(context.Web);
            context.Load(context.Web.Lists);
            context.Load(context.Web, web => web.ServerRelativeUrl);
            context.ExecuteQuery();
            WriteMessage($"Connected to: {context.Web.ServerRelativeUrl}", ConsoleColor.Green);

            List list = context.Web.Lists.Single(l => l.Title == "Documents");
            context.Load(list);
            context.ExecuteQuery();
            WriteMessage($"Number of files: {list.ItemCount}", ConsoleColor.Green);
            
            foreach (var folderToClean in folderRootRelativeUrls)
            {
                var folderUrls = GetFolderList(context, list, folderToClean, recurse);

                foreach (string rootRelativeUrl in folderUrls)
                {
                    WriteMessage($"\r\nFolder: {rootRelativeUrl}", ConsoleColor.Green);
                    ProcessFolder(context, list, rootRelativeUrl);
                }
            }

            WriteMessage("\r\nDone.", ConsoleColor.Green);
            WriteMessage("\r\nPress any key to exit.", ConsoleColor.Green);
            Console.ReadKey();
        }

        private static List<string> GetFolderList(ClientContext context, List list, string rootRelativeUrl, bool recurse = false, List<string> allFolders = null)
        {
            if (allFolders == null)
            {
                allFolders = new List<string>();
            }
            
            Folder folder = context.Web.GetFolderByServerRelativePath(ResourcePath.FromDecodedUrl(context.Web.ServerRelativeUrl + rootRelativeUrl));
            context.Load(folder);
            context.ExecuteQuery();

            allFolders.Add(folder.ServerRelativeUrl);

            if (recurse)
            {
                context.Load(folder.Folders);
                context.ExecuteQuery();

                foreach (var subfolder in folder.Folders)
                {
                    WriteMessage(rootRelativeUrl + '/' + subfolder.Name, ConsoleColor.Yellow);
                    GetFolderList(context, list, rootRelativeUrl + '/' + subfolder.Name, recurse, allFolders);
                }
            }

            return allFolders;
        }

        private static void ProcessFolder(ClientContext context, List list, string rootRelativeUrl)
        {
            const int pageSize = 100;

            Folder folder = context.Web.GetFolderByServerRelativePath(ResourcePath.FromDecodedUrl(rootRelativeUrl));
            context.Load(folder);
            context.ExecuteQuery();

            CamlQuery query = new CamlQuery
            {
                ViewXml = $@"<View>
                       <RowLimit>{pageSize}</RowLimit>
                       <Query>
                         <Where>
                           <Eq>
                             <FieldRef Name='ContentType'/>
                             <Value Type='Computed'>Document</Value>
                           </Eq>
                         </Where>
                       </Query>
                     </View>",
                FolderServerRelativeUrl = folder.ServerRelativeUrl
            };

            bool hasMoreRecords = false;
            int pageCount = 1;

            do
            {
                WriteMessage($"\r\nPage: {pageCount}", ConsoleColor.White);
                ListItemCollection items = list.GetItems(query);
                context.Load(items);
                context.ExecuteQuery();

                ProcessItems(context, items);

                hasMoreRecords = items.ListItemCollectionPosition != null;
                query.ListItemCollectionPosition = items.ListItemCollectionPosition;

                pageCount++;
            } while (hasMoreRecords);
        }

        private static void ProcessItems(ClientContext context, ListItemCollection items)
        {
            foreach (ListItem item in items)
            {
                context.Load(item);
                context.ExecuteQuery();
                ProcessFile(context, item);
            }
        }

        private static void ProcessFile(ClientContext context, ListItem item)
        {
            File file = item.File;

            if (file != null)
            {
                context.Load(file);
                context.Load(file.Versions);
                context.ExecuteQuery();
                long fileSize = file.Length;
                int versionCount = file.Versions.Count;

                if (versionCount > 0)
                {
                    WriteMessage($"File: {file.Name,30}, Version count: {versionCount,5}, Size: {fileSize,8}, Deleting versions...", ConsoleColor.Gray);
                    file.Versions.DeleteAll();
                }
                else
                {
                    WriteMessage($"File: {file.Name,30}, Version count: {versionCount,5}", ConsoleColor.DarkGray);
                }
            }
        }

        private static void WriteMessage(string message, ConsoleColor color)
        {
            Console.ForegroundColor = color;
            Console.WriteLine(message);
            Console.ResetColor();
        }
    }
}
