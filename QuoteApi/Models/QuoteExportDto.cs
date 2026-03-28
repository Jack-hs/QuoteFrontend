public class QuoteExportDto
{
    public BasicInfo basicInfo { get; set; }
    public List<FeeItem> CourseFees { get; set; } = new();
    public List<FeeItem> LocalFees { get; set; } = new();
    public List<FeeItem> OtherFees { get; set; } = new();
    public TotalsDto Totals { get; set; }
}

public class BasicInfo
{
    public string studentName { get; set; }
    public string school { get; set; }
    public string course { get; set; }
    public string roomType { get; set; }
    public string placeOfStay { get; set; }
    public string startDate { get; set; }
    public string endDate { get; set; }
    public string weeks { get; set; }
    public double usdRate { get; set; }
    public double phpRate { get; set; }
    public string airTicket { get; set; }
    public string visa { get; set; }
    public string insurance { get; set; }
    // ... 其他欄位
}

public class TotalsDto
{
    public decimal currentTotalUSD { get; set; }
    public decimal currentTotalNTD { get; set; }
    public decimal currentLocalTotalPeso { get; set; }
    public decimal currentLocalTotalNTD { get; set; }
    public decimal currentOtherTotalNTD { get; set; }
    public decimal AllTotalNTD { get; set; }
    // ... 你的 totals 欄位
}
public class FeeItem
{
    public string key { get; set; } = "";
    public string item { get; set; }
    public string content { get; set; }
    public string weeks { get; set; }
    public string people { get; set; }
    public decimal? unitPrice { get; set; }
    public decimal? amount { get; set; }
    public string remark { get; set; }
}

