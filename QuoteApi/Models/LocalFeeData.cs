using System;
using System.Collections.Generic;
using System.Text;

namespace QuoteApi.Models
{
    public class LocalFeeData
    {
        public class FeeByWeek : Dictionary<string, decimal> { }

        // 第 2 層：雜費項目名稱 -> 每週對應的價格字典 (例如 "SSP ACR E-CARD" -> FeeByWeek)
        public class FeeItems : Dictionary<string, FeeByWeek> { }

        // 第 1 層 (根目錄)：大分類 -> 雜費項目列表 (例如 "DORMITORY" -> FeeItems)
        public class LocalFeeRoot : Dictionary<string, FeeItems> { }
    }
}
