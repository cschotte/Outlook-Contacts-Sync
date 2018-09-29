using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NavaTron.Outlook.Contacts.Sync.Models
{
    class UserDetailsModel
    {
        public int Id { get; set; }

        public string MobileTelephoneNumber { get; set; }

        public string CompanyName { get; set; }

        public string OfficeLocation { get; set; }

        public string FirstName { get; set; }

        public string ManagerName { get; set; }

        public string Account { get; set; }

        public string BusinessTelephoneNumber { get; set; }

        public string Department { get; set; }

        public string FullName { get; set; }

        public string EmailAddress { get; set; }

        public string JobTitle { get; set; }

        public string LastName { get; set; }

        public string NickName { get; set; }

        public byte[] Picture { get; set; }
    }
}
