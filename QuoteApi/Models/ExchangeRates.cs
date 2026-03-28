public class ExchangeRates
{
    public DateTime LastUpdated { get; set; }
    public CurrencyRate USD { get; set; }
    public CurrencyRate PHP { get; set; }
}

public class CurrencyRate
{
    public string Bank { get; set; }
    public decimal BuyRate { get; set; }
    public decimal SellRate { get; set; }
}