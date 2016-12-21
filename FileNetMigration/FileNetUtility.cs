using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;
using System.Security;
using System.Collections;
using MimeTypes;
using Microsoft.SharePoint.Client.Utilities;
using System.Globalization;

namespace FileNetMigration
{
    public class FileNetUtility
    {
        public User currentUser = null;
        public List<List<DocumentObject>> ReadJSON()
        {
            string strpath = @"D:\FileNet\";
            string contents = string.Empty;
            
            List<List<DocumentObject>> JsonData = new List<List<DocumentObject>>();

            string[] files = Directory.GetFiles(strpath, "*.json");
            foreach (string file in files)
            {
                List<DocumentObject> documentData = new List<DocumentObject>();
                contents = System.IO.File.ReadAllText(file);

                //string contents = System.IO.File.ReadAllText(@"D:\FileNet\{A0DD2955-0000-C6F5-9365-DDFB26853F44}.json");

                JObject results = JObject.Parse(contents);

                foreach (var result in results["Documents"])
                {
                    DocumentObject docFile = new DocumentObject();
                    docFile.DocId = (string)result["DocId"];
                    docFile.LocalFileName = (string)result["LocalFileName"];
                    docFile.Owner = (string)result["Owner"];
                    docFile.FileName = (string)result["FileName"];
                    docFile.MimeType = (string)result["MimeType"];
                    JToken versionTypes = result["Version"];
                   
                    docFile.MajorVersion = Convert.ToInt32(((JValue)(versionTypes["MajorVersion"])).Value.ToString());
                    docFile.MinorVersion = Convert.ToInt32(((JValue)(versionTypes["MinorVersion"])).Value.ToString());

                    JToken JsonProperties = result["Properties"];

                    //DocumentProperties props = new DocumentProperties();
                    
                    foreach (JObject item in JsonProperties)
                    {
                        if (item.Property("Creator") != null)
                            docFile.Creator = ((JProperty)(item["Creator"].First)).Name;
                        if (item.Property("LastModifier") != null)
                            docFile.LastModifier = ((JProperty)(item["LastModifier"].First)).Name;
                        if (item.Property("DateCreated") != null)
                            docFile.DateCreated =((JProperty)(item["DateCreated"].First)).Name;
                        if (item.Property("DateLastModified") != null)
                            docFile.DateLastModified = ((JProperty)(item["DateLastModified"].First)).Name;
                    }
                    docFile.VersionStatus = result["VersionStatus"].ToString();
                    
                    docFile.FiledInFolder = @"D:\FileNet\" + docFile.LocalFileName;
                    docFile.DocumentName = result["DocumentName"].ToString();
                    if (result["LatestVersion"] != null)
                        docFile.LatestVersion = result["LatestVersion"].ToString();
                    else
                        docFile.LatestVersion = null;
                    documentData.Add(docFile);

                }

                documentData = documentData.OrderBy(e => e.MajorVersion).ThenBy(e => e.MinorVersion).ToList();

                JsonData.Add(documentData);
            }
            Console.WriteLine(JsonData.Count());
            return JsonData;
        }
        private SecureString GetPasswordFromConsoleInput()
        {
            ConsoleKeyInfo info;
            //Get the user's password as a SecureString
            SecureString securePassword = new SecureString();
            do
            {
                info = Console.ReadKey(true);
                if (info.Key != ConsoleKey.Enter)
                {
                    securePassword.AppendChar(info.KeyChar);
                }
            }
            while (info.Key != ConsoleKey.Enter);
            return securePassword;
        }

        public ClientContext ConnectionSharePointOnline()
        {
            string webUrl = "https://pepsico.sharepoint.com/sites/GlobalInsights";
            string userName = "sptest1.sptest1@pepsico.com";
            string password = "pass@word35";


            try
            {
                ConsoleColor defaultForeground = Console.ForegroundColor;

                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("Enter the URL of the SharePoint Online site:");

                Console.ForegroundColor = defaultForeground;
                //string webUrl = Console.ReadLine();

                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("Enter your user name (ex: test@mytenant.microsoftonline.com):");
                Console.ForegroundColor = defaultForeground;
                //string userName = Console.ReadLine();

                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("Enter your password.");
                Console.ForegroundColor = defaultForeground;
                //SecureString securePassword = GetPasswordFromConsoleInput();
                SecureString securePassword = new SecureString();
                foreach (char ch in password.ToCharArray())
                {
                    securePassword.AppendChar(ch);
                }

                using (var context = new ClientContext(webUrl))
                {
                    context.Credentials = new SharePointOnlineCredentials(userName, securePassword);
                    context.Load(context.Web, w => w.Title);
                    context.ExecuteQuery();

                    Console.ForegroundColor = ConsoleColor.White;
                    Console.WriteLine("Your site title is: " + context.Web.Title);

                    currentUser = context.Web.EnsureUser(userName);
                    context.Load(currentUser);

                    return context;
                }

            }
            catch (Exception ex)
            {
                throw ex;
            }

        }

        public void UploadDocumentSP(ClientContext context, DocumentObject doc)
        {
            User createdBy, modifiedBy, serviceAccount = null;
            try
            {
                FileCreationInformation fileCreationInfo = null;
                using (FileStream fs = new FileStream(@"D:\FileNet\" + doc.LocalFileName, FileMode.Open))
                {
                    fileCreationInfo = new FileCreationInformation();
                    fileCreationInfo.ContentStream = fs;
                    fileCreationInfo.Overwrite = true;
                    fileCreationInfo.Url = doc.FileName;

                    List docs = context.Web.Lists.GetByTitle("FileNetDocuments");
                    context.Load(docs.RootFolder);
                    context.ExecuteQuery();
                    Folder NewFolder = EnsureFolder(context, docs.RootFolder, "/Folder1/Folder2/Folder3");

                    Microsoft.SharePoint.Client.File uploadFile = NewFolder.Files.Add(fileCreationInfo);
                    if (doc.MinorVersion > 0)
                    {
                        uploadFile.CheckOut();

                        uploadFile.CheckIn("Minor Version", CheckinType.MinorCheckIn);
                    }
                    //context.Load(uploadFile);
                    //context.ExecuteQuery();

                    docs.EnableVersioning = false;
                    docs.Update();
                    context.ExecuteQuery();

                    ListItem item = uploadFile.ListItemAllFields;
                    context.Load(item);
                    createdBy = CheckUserExist(context, Convert.ToString(doc.Owner));
                    modifiedBy = CheckUserExist(context, Convert.ToString(doc.LastModifier));
                    //Updating Metadata Created, CreatedBy, Modified, Modified By
                    if (createdBy != null && modifiedBy != null)
                    {
                        item["Author"] = (createdBy.Id + ";#" + createdBy.LoginName);
                        item["Editor"] = (modifiedBy.Id + ";#" + modifiedBy.LoginName);
                    }
                    else
                    {
                        serviceAccount = CheckUserExist(context, Convert.ToString("Jagadish.Subramonayan.Contractor@pepsico.com"));
                        item["Author"] = (serviceAccount.Id + ";#" + serviceAccount.LoginName);
                        docs.Fields.GetByInternalNameOrTitle("Editor").ReadOnlyField = false;
                        item["Editor"] = (serviceAccount.Id + ";#" + serviceAccount.LoginName);
                    }
                    DateTime createdDate = DateTime.ParseExact(doc.DateCreated, "ddd MMM dd HH:mm:ss IST yyyy", new CultureInfo("en-us"));
                    DateTime modifiedDate = DateTime.ParseExact(doc.DateLastModified, "ddd MMM dd HH:mm:ss IST yyyy", new CultureInfo("en-us"));
                    item["Created"] = createdDate;
                    item["Modified"] = modifiedDate;
                    item.Update();
                    //uploadFile.CheckOut();
                    //uploadFile.CheckIn("Initial Version", CheckinType.OverwriteCheckIn);
                    //context.ExecuteQuery();
                    //Enable Version settings for DOC LIB after update Metadata
                    docs.EnableVersioning = true;
                    docs.Update();
                    context.ExecuteQuery();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
        //public bool TryGetFileByServerRelativeUrl(ClientContext context, string serverRelativeUrl)
        //{
        //    try
        //    {
        //        Microsoft.SharePoint.Client.File file = context.Web.GetFileByServerRelativeUrl(serverRelativeUrl);
        //        context.Load(file);
        //        context.ExecuteQuery();
        //        return true;
        //    }
        //    catch (Microsoft.SharePoint.Client.ServerException ex)
        //    {
        //        Console.ForegroundColor = ConsoleColor.Red;
        //        Console.WriteLine(ex.Message);
        //        if (ex.ServerErrorTypeName == "System.IO.FileNotFoundException")
        //        {
        //            Console.ForegroundColor = ConsoleColor.Red;
        //            Console.WriteLine(ex.Message);
        //            return false;
        //        }
        //    }

        //}

        public  Folder EnsureFolder(ClientContext ctx, Folder ParentFolder, string FolderPath)
        {
            //Split up the incoming path so we have the first element as the a new sub-folder name 
            //and add it to ParentFolder folders collection
            string[] PathElements = FolderPath.Split(new char[] { '/' }, StringSplitOptions.RemoveEmptyEntries);
            string Head = PathElements[0];
            Folder NewFolder = ParentFolder.Folders.Add(Head);
            ctx.Load(NewFolder);
            ctx.ExecuteQuery();

            //If we have subfolders to create then the length of PathElements will be greater than 1
            if (PathElements.Length > 1)
            {
                //If we have more nested folders to create then reassemble the folder path using what we have left i.e. the tail
                string Tail = string.Empty;
                for (int i = 1; i < PathElements.Length; i++)
                    Tail = Tail + "/" + PathElements[i];

                //Then make a recursive call to create the next subfolder
                return EnsureFolder(ctx, NewFolder, Tail);
            }
            else
                //This ensures that the folder at the end of the chain gets returned
                return NewFolder;
        }
        public void UploadDocument(ClientContext context, List<List<DocumentObject>> documentsList)
        {
          
            if (documentsList.Count > 0)
            {
                foreach (List<DocumentObject> document in documentsList)
                {
                    foreach (DocumentObject doc in document)
                    {
                        if (!string.IsNullOrEmpty(doc.FileName) && !string.IsNullOrEmpty(doc.LocalFileName))
                        {
                            string uploadDocPath = doc.FiledInFolder;
                            string docExtn = Path.GetExtension(uploadDocPath);

                            if (!string.IsNullOrEmpty(docExtn))
                            {                               
                                UploadDocumentSP(context, doc);
                            }
                            else
                            {
                                string strMimetype = MimeTypeMap.GetExtension(doc.MimeType);
                                doc.FileName = doc.FileName + strMimetype;
                                UploadDocumentSP(context, doc);
                            }

                        }
                    }
                }
            }
        }
        private User CheckUserExist(ClientContext context, string userID)
        {
            //userID = "sptest2.sptest2@pepsico.com";
            User user = null;
            try
            {
                var result = Microsoft.SharePoint.Client.Utilities.Utility.ResolvePrincipal(context, context.Web, userID, Microsoft.SharePoint.Client.Utilities.PrincipalType.User, Microsoft.SharePoint.Client.Utilities.PrincipalSource.All, null, true);
                context.ExecuteQuery();
                if (result != null)
                {
                    user = context.Web.EnsureUser(result.Value.LoginName);
                    context.Load(user);
                    context.ExecuteQuery();
                }
               
            }
            catch (Exception ex)
            {
                Console.WriteLine("User Exception: " + ex.Message);
                //user = currentUser;
                user = null;
            }
            return user;
        }
    }

}
