using System;
using System.Collections.Generic;
using System.Text;

namespace QuoteApi.Models
{
    public class TuitionData
    {
        // 第 4 層：週數 -> 價格 (例如 "1W" -> 1170)
        public class DurationPricing : Dictionary<string, decimal> { }

        // 第 3 層：房型 -> 週數與價格 (例如 "Single" -> DurationPricing)
        public class RoomPricing : Dictionary<string, DurationPricing> { }

        // 第 2 層：課程名稱 -> 房型與價格 (例如 "IELTS INTENSIVE" -> RoomPricing)
        public class CourseCategory : Dictionary<string, RoomPricing> { }

        // 第 1 層 (根目錄)：大分類 -> 課程詳細資料 (例如 "IELTS" -> CourseCategory)
        // 整個檔案反序列化後就是這個類別
        public class TuitionRoot : Dictionary<string, CourseCategory> { }
    }
}
