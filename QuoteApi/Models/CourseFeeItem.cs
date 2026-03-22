using System;
using System.Collections.Generic;
using System.Text;

namespace QuoteApi.Models
{
    public class CourseFeeItem
    {
        public string Key { get; set; }
        public string Item { get; set; }
        public string Content { get; set; }
        public string Weeks { get; set; }
        public int People { get; set; }
        public string UnitPrice { get; set; }
        public decimal Amount { get; set; }
        public string Remark { get; set; }
    }
}
