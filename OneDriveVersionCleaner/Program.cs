using System;
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
      string webFullUrl = "https://TODO.sharepoint.com/personal/TODO_onmicrosoft_com";
      string username = "TODO@TODO.onmicrosoft.com";
      SecureString password = new NetworkCredential("", "TODO-ENTER-YOUR-APP-PASSWORD-HERE").SecurePassword;  // HACK :)

      string[] folderRootRelativeUrls = {
          "/Documents/subfolder1",
          "/Documents/subfolder1/subfolder11",
          "/Documents/subfolder1/subfolder11/subfolder111",
          "/Documents/subfolder1/subfolder12",
          "/Documents/subfolder2/subfolder21"
      };

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

      foreach (string rootRelativeUrl in folderRootRelativeUrls)
      {
        WriteMessage($"\r\nFolder: {rootRelativeUrl}", ConsoleColor.Green);
        ProcessFolder(context, list, rootRelativeUrl);
      }

      WriteMessage("\r\nDone.", ConsoleColor.Green);
    }

    private static void ProcessFolder(ClientContext context, List list, string rootRelativeUrl)
    {
      const int pageSize = 100;

      Folder folder = context.Web.GetFolderByServerRelativeUrl(context.Web.ServerRelativeUrl + rootRelativeUrl);
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
          WriteMessage($"File: {file.Name,30}, Version count: {versionCount,5}, Size: {fileSize, 8}, Deleting versions...", ConsoleColor.Gray);
          // TODO: file.Versions.DeleteAll();
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
