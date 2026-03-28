using Microsoft.Extensions.Caching.Memory;
using System;
using System.Collections.Generic;
using System.Text;
using static QuoteApi.Models.TuitionData;

namespace QuoteApi.Models
{
    public static class ZoneData
    {
        public static string StartupPath = "";
        public static TuitionRoot tuitionData = new TuitionRoot();
        public static AppSettings appSettings = new AppSettings();
        public static Dictionary<string, List<Tuple<int, int>>> QuotationTermsDict = new Dictionary<string, List<Tuple<int, int>>>();
    }
}
