# OneDrive / SharePoint Version Cleaner

This sample console application uses the SharePoint .NET client-side object model (CSOM) to delete old versions of all files from multiple folders from your SharePoint or OneDrive for Business subscription.

## Important!

**This is not production code, and it deletes files from your cloud storage. Use at your own risk!**

## How it works

This console application executes the following steps:

1. Connects to your OneDrive for Business or SharePoint site collection.
2. Finds the `Documents` document library.
3. Iterates through the specified subfolder paths in the document library.
4. It runs CAML queries in every folder to retrieve the documents (files). In a single query maximum 100 documents are retrieved, and the query is executed again and again until all documents are processed.
5. If a document has multiple versions, they are deleted.

## How to use it

To use this sample follow these steps:

1. Clone this repository.
2. Open the `OneDriveVersionCleanerSolution.sln` solution in Visual Studio (written and tested with VS 2017).
3. At the top of the `Main` method specify the URL of your cloud storage, your username and password, and the paths to the folders you want to process. If you have multifactor authentication enabled (you do, right?) you need to create an application password first (check the blog post mentioned in the Read more section below).
4. In line 124 uncomment the `file.Versions.DeleteAll()` call to actually delete the old versions.
5. Run at your own risk.

## Read more

You can read the full story behind this code and other options to reduce the OneDrive storage space needed for your files on my [Delete old document versions from OneDrive for Business blog post](https://gyorgybalassy.wordpress.com/2018/11/11/delete-old-document-versions-from-sharepoint-onedrive/). 

## About the author

This sample was created by [György Balássy](https://linkedin.com/in/balassy).
