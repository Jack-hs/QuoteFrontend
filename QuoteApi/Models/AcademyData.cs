using System;
using System.Collections.Generic;
using System.Text;
using System.Text.Json.Serialization;

namespace QuoteApi.Models
{
    public class AcademyData
    {
        public DateTime StartDate;
        public DateTime EndDate;
        public string RoomTypeTranslate;
        public string FlightDest;
        public string SchoolLocation;
        public int registrationfee;
        public bool registrationfeeInclusive;
        public string Name;
        public string DurWeeks;
        public string CourseFee;
        public int PicLocationRow;
        public int PicLocationCol;
        public int PicLocationLeft;
        public int PicLocationTop;
        public int LogoWidth;
        public int LogoHeight;
        public decimal SummerSurcharge;
        public DateTime SummerSurchargeStartTime;
        public DateTime SummerSurchargeEndTime;
        public decimal DiscountInclusive;
        public decimal Discount;
        public decimal FixedDiscount;
        public int RowHeight = 25;
        public int FontSize = 12;
        public int LegalGuardianFee;
    }

    // 根物件
    public class AppSettings
    {
        [JsonPropertyName("setting")]
        public SettingInfo Setting { get; set; }

        [JsonPropertyName("courseFee")]
        public CourseFeeInfo CourseFee { get; set; }

        [JsonPropertyName("logo")]
        public LogoInfo Logo { get; set; }

        [JsonPropertyName("localFee")]
        public List<LocalFeeInfo> LocalFee { get; set; }

        // 使用 Dictionary 來對應動態的分類名稱 ("IELTS", "ESL & SPEAKING" 等)
        [JsonPropertyName("courses")]
        public Dictionary<string, List<CourseInfo>> Courses { get; set; }

        [JsonPropertyName("room")]
        public List<RoomInfo> Room { get; set; }

        [JsonPropertyName("more4week")]
        public Dictionary<string, string> More4Week { get; set; }

        [JsonPropertyName("less4week")]
        public Dictionary<string, string> Less4Week { get; set; }
    }

    public class SettingInfo
    {
        [JsonPropertyName("schoolLocation")]
        public string SchoolLocation { get; set; }

        [JsonPropertyName("flightDest")]
        public string FlightDest { get; set; }
    }

    public class CourseFeeInfo
    {
        [JsonPropertyName("registrationfee")]
        public string RegistrationFee { get; set; }

        [JsonPropertyName("registrationfeeInclusive")]
        public string RegistrationFeeInclusive { get; set; }

        [JsonPropertyName("summerSurcharge")]
        public string SummerSurcharge { get; set; }

        [JsonPropertyName("summerSurchargeStartTime")]
        public string SummerSurchargeStartTime { get; set; }

        [JsonPropertyName("summerSurchargeEndTime")]
        public string SummerSurchargeEndTime { get; set; }

        [JsonPropertyName("discountInclusive")]
        public string DiscountInclusive { get; set; }

        [JsonPropertyName("discount")]
        public string Discount { get; set; }

        [JsonPropertyName("fixedDiscount")]
        public string FixedDiscount { get; set; }

        [JsonPropertyName("legalGuardianFee")]
        public string LegalGuardianFee { get; set; }
    }

    public class LogoInfo
    {
        [JsonPropertyName("picLocationRow")]
        public string PicLocationRow { get; set; }

        [JsonPropertyName("picLocationCol")]
        public string PicLocationCol { get; set; }

        [JsonPropertyName("picLocationLeft")]
        public string PicLocationLeft { get; set; }

        [JsonPropertyName("picLocationTop")]
        public string PicLocationTop { get; set; }

        [JsonPropertyName("picwidth")]
        public string PicWidth { get; set; }

        [JsonPropertyName("picheight")]
        public string PicHeight { get; set; }
    }

    public class LocalFeeInfo
    {
        [JsonPropertyName("code")]
        public string Code { get; set; }

        [JsonPropertyName("DocItem")]
        public string DocItem { get; set; }

        [JsonPropertyName("Item")]
        public string Item { get; set; }

        [JsonPropertyName("content")]
        public string Content { get; set; }

        [JsonPropertyName("times")]
        public string Times { get; set; }

        [JsonPropertyName("price")]
        public string Price { get; set; }

        [JsonPropertyName("remark ")] // 注意 JSON 這裡有一個空格！
        public string Remark { get; set; }
    }

    public class CourseInfo
    {
        [JsonPropertyName("name")]
        public string Name { get; set; }

        [JsonPropertyName("code")]
        public string Code { get; set; }

        [JsonPropertyName("pricePerWeek")]
        public decimal PricePerWeek { get; set; }

        [JsonPropertyName("description")]
        public string Description { get; set; }
    }

    public class RoomInfo
    {
        [JsonPropertyName("Name")]
        public string Name { get; set; }

        [JsonPropertyName("roomType")]
        public string RoomType { get; set; }

        [JsonPropertyName("code")]
        public string Code { get; set; }

        [JsonPropertyName("pricePerWeek")]
        public decimal PricePerWeek { get; set; }

        [JsonPropertyName("description")]
        public string Description { get; set; }
    }
}
