using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.DirectoryServices;
using System.Diagnostics;
using System.Runtime.InteropServices;
using NavaTron.Outlook.Contacts.Sync.Models;
using System.IO;
using System.Text.RegularExpressions;

namespace NavaTron.Outlook.Contacts.Sync.Controllers
{
    class SyncController
    {
        private readonly Microsoft.Office.Interop.Outlook.Application outlookApp;
        private readonly List<UserDetailsModel> domainUsers = new List<UserDetailsModel>();
        private readonly List<UserDetailsModel> outlookUsers = new List<UserDetailsModel>();

        public SyncController()
        {
            outlookApp = GetApplicationObject();
        }

        private Microsoft.Office.Interop.Outlook.Application GetApplicationObject()
        {
            Microsoft.Office.Interop.Outlook.Application app;
            Process[] processes = Process.GetProcessesByName("OUTLOOK");

            if (processes.Length > 0)
            {
                try
                {
                    app = (Microsoft.Office.Interop.Outlook.Application)Marshal.GetActiveObject("Outlook.Application");

                    return app;
                }
                catch { }
            }

            app = new Microsoft.Office.Interop.Outlook.Application();

            return app;
        }

        internal void GetDomainUsers()
        {
            // Get Doamin users
            DirectorySearcher searcher = new DirectorySearcher
            {
                Filter = Properties.Settings.Default.Filter,
                SizeLimit = 1500,
                PageSize = 1500
            };

            searcher.PropertiesToLoad.Add("mobile");
            searcher.PropertiesToLoad.Add("company");
            searcher.PropertiesToLoad.Add("physicaldeliveryofficename");
            searcher.PropertiesToLoad.Add("givenname");
            searcher.PropertiesToLoad.Add("manager");
            searcher.PropertiesToLoad.Add("samaccountname");
            searcher.PropertiesToLoad.Add("telephonenumber");
            searcher.PropertiesToLoad.Add("thumbnailphoto");
            searcher.PropertiesToLoad.Add("department");
            searcher.PropertiesToLoad.Add("displayname");
            searcher.PropertiesToLoad.Add("mail");
            searcher.PropertiesToLoad.Add("title");
            searcher.PropertiesToLoad.Add("sn");

            SearchResultCollection users = searcher.FindAll();

            // Add Doamin users to List
            domainUsers.Clear();

            for (int i = 0; i < users.Count; i++)
            {
                var item = new UserDetailsModel();

                if (users[i].Properties.Contains("mobile")) item.MobileTelephoneNumber = users[i].Properties["mobile"][0].ToString();
                if (users[i].Properties.Contains("company")) item.CompanyName = users[i].Properties["company"][0].ToString();
                if (users[i].Properties.Contains("physicaldeliveryofficename")) item.OfficeLocation = users[i].Properties["physicaldeliveryofficename"][0].ToString();
                if (users[i].Properties.Contains("givenname")) item.FirstName = users[i].Properties["givenname"][0].ToString();
                if (users[i].Properties.Contains("manager")) item.ManagerName = users[i].Properties["manager"][0].ToString();
                if (users[i].Properties.Contains("samaccountname")) item.Account = users[i].Properties["samaccountname"][0].ToString();
                if (users[i].Properties.Contains("telephonenumber")) item.BusinessTelephoneNumber = users[i].Properties["telephonenumber"][0].ToString();
                if (users[i].Properties.Contains("department")) item.Department = users[i].Properties["department"][0].ToString();
                if (users[i].Properties.Contains("displayname")) item.FullName = users[i].Properties["displayname"][0].ToString();
                if (users[i].Properties.Contains("mail")) item.EmailAddress = users[i].Properties["mail"][0].ToString();
                if (users[i].Properties.Contains("title")) item.JobTitle = users[i].Properties["title"][0].ToString();
                if (users[i].Properties.Contains("sn")) item.LastName = users[i].Properties["sn"][0].ToString();
                if (users[i].Properties.Contains("samaccountname")) item.NickName = users[i].Properties["samaccountname"][0].ToString();
                if (users[i].Properties.Contains("thumbnailphoto")) item.Picture = (byte[])users[i].Properties["thumbnailphoto"][0];

                domainUsers.Add(item);
            }
        }

        internal void GetOutlookUsers()
        {
            Microsoft.Office.Interop.Outlook.MAPIFolder folder = outlookApp.Session.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderContacts);

            outlookUsers.Clear();

            for (int i = folder.Items.Count; i > 0; i--)
            {
                try
                {
                    Microsoft.Office.Interop.Outlook.ContactItem user = folder.Items[i];

                    if (user.Categories == Properties.Settings.Default.Categories)
                    {
                        var item = new UserDetailsModel
                        {
                            Id = i,
                            MobileTelephoneNumber = user.MobileTelephoneNumber,
                            CompanyName = user.CompanyName,
                            OfficeLocation = user.OfficeLocation,
                            FirstName = user.FirstName,
                            ManagerName = user.ManagerName,
                            Account = user.Account,
                            BusinessTelephoneNumber = user.BusinessTelephoneNumber,
                            Department = user.Department,
                            FullName = user.FullName,
                            EmailAddress = user.Email1Address,
                            JobTitle = user.JobTitle,
                            LastName = user.LastName,
                            NickName = user.NickName,
                        };

                        outlookUsers.Add(item);
                    }
                }
                catch { }
            }
        }

        internal void UpdateOutlookUsers()
        {
            var updateUsers = from o in outlookUsers
                              from d in domainUsers
                              where o.NickName == d.NickName
                              select new { o, d };

            Microsoft.Office.Interop.Outlook.MAPIFolder folder = outlookApp.Session.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderContacts);

            foreach (var user in updateUsers)
            {
                Microsoft.Office.Interop.Outlook.ContactItem item = folder.Items[user.o.Id];

                item.MobileTelephoneNumber = user.d.MobileTelephoneNumber;
                item.CompanyName = user.d.CompanyName;
                item.OfficeLocation = user.d.OfficeLocation;
                item.FirstName = user.d.FirstName;
                item.ManagerName = user.d.ManagerName;
                item.Account = user.d.Account;
                item.BusinessTelephoneNumber = user.d.BusinessTelephoneNumber;
                item.Department = user.d.Department;
                item.FullName = user.d.FullName;
                item.Email1Address = user.d.EmailAddress;
                item.JobTitle = user.d.JobTitle;
                item.LastName = user.d.LastName;
                item.NickName = user.d.NickName;

                string filename = string.Empty;

                try
                {
                    filename = AddPhoto(user.d.NickName, user.d.Picture);

                    if (!string.IsNullOrEmpty(filename)) item.AddPicture(filename);
                }
                catch { }

                item.Save();

                // cleanup
                if (!string.IsNullOrEmpty(filename)) DeletePhoto(filename);
            }

        }

        internal void RemoveOutlookUsers()
        {
            var oldUsers = from o in outlookUsers
                           where !(from d in domainUsers
                                   select d.NickName)
                                   .Contains(o.NickName)
                           orderby o.Id descending
                           select o;

            Microsoft.Office.Interop.Outlook.MAPIFolder folder = outlookApp.Session.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderContacts);

            foreach (var user in oldUsers)
            {
                folder.Items.Remove(user.Id);
            }
        }

        internal void AddOutLookUsers()
        {
            var newUsers = from d in domainUsers
                           where !(from o in outlookUsers
                                   select o.NickName)
                                   .Contains(d.NickName)
                           select d;

            Microsoft.Office.Interop.Outlook.MAPIFolder folder = outlookApp.Session.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderContacts);

            foreach (var user in newUsers)
            {
                Microsoft.Office.Interop.Outlook.ContactItem item = (Microsoft.Office.Interop.Outlook.ContactItem)folder.Items.Add(Microsoft.Office.Interop.Outlook.OlItemType.olContactItem);

                item.MobileTelephoneNumber = user.MobileTelephoneNumber;
                item.CompanyName = user.CompanyName;
                item.OfficeLocation = user.OfficeLocation;
                item.FirstName = user.FirstName;
                item.ManagerName = user.ManagerName;
                item.Account = user.Account;
                item.BusinessTelephoneNumber = user.BusinessTelephoneNumber;
                item.Department = user.Department;
                item.FullName = user.FullName;
                item.Email1Address = user.EmailAddress;
                item.JobTitle = user.JobTitle;
                item.LastName = user.LastName;
                item.NickName = user.NickName;


                string filename = string.Empty;

                try
                {
                    filename = AddPhoto(user.NickName, user.Picture);

                    if (!string.IsNullOrEmpty(filename)) item.AddPicture(filename);
                }
                catch { }

                item.Categories = Properties.Settings.Default.Categories;
                item.Save();

                // Cleanup
                if (!string.IsNullOrEmpty(filename)) DeletePhoto(filename);
            }
        }

        private string AddPhoto(string alias, byte[] data)
        {
            if (data == null) return string.Empty;

            try
            {
                string filename = Environment.GetEnvironmentVariable("TEMP") + @"\" + alias + ".jpg";

                using (var writeStream = new FileStream(filename, FileMode.Create))
                {
                    using (var writeBinary = new BinaryWriter(writeStream))
                    {
                        writeBinary.Write(data);
                        writeBinary.Close();
                    }
                }

                return filename;
            }
            catch
            {
                return string.Empty;
            }
        }

        private void DeletePhoto(string filename)
        {
            if (File.Exists(filename)) File.Delete(filename);
        }

        internal void FixUserDetails()
        {
            for (int i = 0; i < domainUsers.Count; i++)
            {
                // fix MobileTelephoneNumber
                if (!string.IsNullOrWhiteSpace(domainUsers[i].MobileTelephoneNumber))
                {
                    string num = Regex.Replace(domainUsers[i].MobileTelephoneNumber, @"\D", "");

                    domainUsers[i].MobileTelephoneNumber = "+" + num;
                }

                // fix BusinessTelephoneNumber
                if (!string.IsNullOrWhiteSpace(domainUsers[i].BusinessTelephoneNumber))
                {
                    string tel = domainUsers[i].BusinessTelephoneNumber;

                    domainUsers[i].BusinessTelephoneNumber = tel.Split('X')[0];
                }

                // fix ManagerName
                if (!string.IsNullOrWhiteSpace(domainUsers[i].ManagerName))
                {
                    Match match = Regex.Match(domainUsers[i].ManagerName, @"CN=([A-Za-z0-9\- ]+)\,");

                    if (match.Success)
                    {
                        domainUsers[i].ManagerName = match.Value.TrimEnd(',').Replace("CN=", "");
                    }
                }
            }
        }
    }
}
