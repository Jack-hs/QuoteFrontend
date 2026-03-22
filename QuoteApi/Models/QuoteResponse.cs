
public class QuoteResponse
{
    public decimal Tuition { get; set; }
    public decimal Dormitory { get; set; }
    public decimal LocalTotal { get; set; }
    public decimal FixedTotal { get; set; }  // 新增這行
    public decimal TotalNT { get; set; }
    public List<LocalFeeItem> LocalFees { get; set; } = new();
}
