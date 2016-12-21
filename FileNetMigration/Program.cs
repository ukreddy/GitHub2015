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
using System.DirectoryServices;

namespace FileNetMigration
{

    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                FileNetUtility objUtility = new FileNetUtility();
                List<List<DocumentObject>> JsonDataObject = objUtility.ReadJSON();
                ClientContext context = objUtility.ConnectionSharePointOnline();
                objUtility.UploadDocument(context, JsonDataObject);

            }
            catch (Exception ex)
            {
                Console.WriteLine("");
                Console.WriteLine(string.Format("Error: {0} \n Stack Trace: {1}", ex.Message, ex.StackTrace));
            }
            //try
            //{
            //    List<Users> lstADUsers = new List<Users>();
            //    //string DomainPath = "LDAP://DC=cts,DC=com";
            //    string DomainPath = "GC://DC=pep,DC=pvt";

            //    DirectorySearcher ds = new DirectorySearcher(DomainPath);
            //    ds.SearchScope = SearchScope.Subtree;
            //    ds.Filter = @"(&(objectClass=user)(SamAccountName=03340687))";
            //    SearchResult rs = ds.FindOne();
            //    Console.WriteLine(rs.GetDirectoryEntry().Properties["displayName"].Value);
            //    Console.WriteLine(rs.GetDirectoryEntry().Properties["mail"].Value);
            //    Console.WriteLine(rs.GetDirectoryEntry().Properties["userAccountControl"].Value);
            //    Console.ReadLine();

            //    //Console.WriteLine(lstADUsers.Count);
            //}
            //catch (Exception ex)
            //{
            //    Console.WriteLine(ex.Message);
            //}
           
            finally
            {
                Console.WriteLine("");
                Console.WriteLine("Done ... Press any key to exit");
                Console.ReadKey();
            }

        }
       
      
    }
    public class Users
    {
        public string Email { get; set; }
        public string UserName { get; set; }
        public string DisplayName { get; set; }
        public bool isMapped { get; set; }
    }

    public class DocumentObject
    {
        public string DocId { get; set; }
        public string LocalFileName { get; set; }
        public string Owner { get; set; }
        public string LatestVersion { get; set; }
        //public List<Version> Version { get; set; }
        public int MajorVersion { get; set; }
        public int MinorVersion { get; set; }
        public string FileName { get; set; }
        //public List<DocumentProperties> DocProperties { get; set; }
        public string FiledInFolder { get; set; }
        public string VersionStatus { get; set; }
        public string MimeType { get; set; }
        public string DocumentName { get; set; }
        public string Creator { get; set; }
        public string LastModifier { get; set; }
        public string DateCreated { get; set; }
        public string DateLastModified { get; set; }
    }

    public class Version
    {
        public string MajorVersion { get; set; }
        public string MinorVersion { get; set; }
    }
    public class DocumentProperties
    {
        public string Creator { get; set; }
        public string LastModifier { get; set; }
        public DateTime DateCreated { get; set; }
        public DateTime DateLastModified { get; set; }

    }
    public class jsonFiles
    {
        public string JsonFile { get; set; }
        public List<DocumentProperties> DocProperties { get; set; }
    }

}
