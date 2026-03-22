// 定義前端傳來的 Request 格式
public class QuoteRequest
{
    public string School { get; set; }
    public string Course { get; set; }
    public string RoomType { get; set; }
    public string Placeofstay { get; set; }
    public int Weeks { get; set; }
    public decimal ExchangeRate { get; set; }
    public bool NeedGuardianFee { get; set; }
    // 👇 日期相關的欄位 👇
    public DateTime StartDate { get; set; } // 預計出發日
    public DateTime EndDate { get; set; }   // 預計結束日
    public DateTime QuoteDate { get; set; } // 報價建立日期 (剛新增的)
    public decimal SchoolDiscount { get; set; } // 👈 新增這個來接收前端輸入的折扣金額

    public decimal AirTicket { get; set; } = 10000;
    public decimal Visa { get; set; } = 1200;
    public decimal Insurance { get; set; } = 1500;
}
//public class QuoteRequest
//{
//    public DateTime StartDate { get; set; }
//    public DateTime EndDate { get; set; }
//    public string School { get; set; } = "PHILINTER";
//    public string Course { get; set; } = "IELTS INTENSIVE";
//    public string RoomType { get; set; } = "Single";
//    public int Weeks { get; set; }
//    public List<LocalFeeItem> LocalFees { get; set; } = new();
//    public decimal AirTicket { get; set; } = 10000;
//    public decimal Visa { get; set; } = 1200;
//    public decimal Insurance { get; set; } = 1500;
//    public decimal ExchangeRate { get; set; } = 32.05m;
//}

public class LocalFeeItem
{
    public string Item { get; set; } = "";
    public decimal Amount { get; set; }
}
