using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Text;
using Commsights.Data.Models;

namespace Commsights.Data.DataTransferObject
{
    public class BaiVietReport
    {
        public string URLCode { get; set; }
        public string URLEmployee { get; set; }
        public string Tier { get; set; }
        public string MediaTitle { get; set; }
        public string Title { get; set; }
        public bool IsDuplicate { get; set; }
        public int? IndustryID { get; set; }
        public string IndustryName { get; set; }
        public int? EmployeeID { get; set; }
        public string Employee { get; set; }
        public int? BaiVietSum { get; set; }
        public int? BaiVietDuplicate { get; set; }
        public int? BaiVietExtant { get; set; }

    }
}
