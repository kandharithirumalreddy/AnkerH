using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace XRMComposeAddinWeb.Models
{
    public class CaseInfo
    {
        public string Title { get; set; }
        public string ID { get; set; }
        //public string Status { get; set; }
        public string FolderPath { get; set; }
        public string Level { get; set; }
        public string CaseFolderName { get; set; }
    }
}