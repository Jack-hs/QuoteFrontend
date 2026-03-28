using System;
using System.Net.Http;
using System.Threading.Tasks;
using System.Linq;
using HtmlAgilityPack;

public class ExchangeRateService
{
    private readonly HttpClient _httpClient;

    public ExchangeRateService(HttpClient httpClient)
    {
        _httpClient = httpClient;
    }

    /// <summary>
    /// 抓取富邦銀行 USD/TWD 匯率 (通常看即期匯率)
    /// </summary>
    public async Task<(string BuySpot, string SellSpot)> GetFubonUsdRateAsync()
    {
        string url = "https://www.fubon.com/Fubon_Portal/banking/Personal/deposit/exchange_rate/exchange_rate1.jsp";

        // 取得網頁 HTML
        var html = await _httpClient.GetStringAsync(url);

        // 使用 HtmlAgilityPack 解析
        var doc = new HtmlDocument();
        doc.LoadHtml(html);

        // 找到所有的表格列 (tr)
        var rows = doc.DocumentNode.SelectNodes("//tr");
        if (rows != null)
        {
            foreach (var row in rows)
            {
                var tds = row.SelectNodes("td");
                // 富邦的結構：tds[1] 是幣別，tds[3] 是現金(買/賣)，tds[4] 是即期(買/賣)
                if (tds != null && tds.Count >= 5)
                {
                    string currencyName = tds[1].InnerText.Trim();
                    if (currencyName.Contains("USD"))
                    {
                        // tds[4] 包含即期買入與賣出，中間有多個空白隔開，我們將其切割
                        string[] spotRates = tds[4].InnerText.Trim()
                            .Split(new char[] { ' ', '\r', '\n', '\t' }, StringSplitOptions.RemoveEmptyEntries);

                        if (spotRates.Length == 2)
                        {
                            return (BuySpot: spotRates[0], SellSpot: spotRates[1]);
                        }
                    }
                }
            }
        }
        throw new Exception("無法解析富邦銀行美金匯率");
    }

    /// <summary>
    /// 抓取台灣銀行 PHP/TWD 匯率 (菲律賓披索通常只看現金匯率)
    /// </summary>
    public async Task<(string BuyCash, string SellCash)> GetBotPhpRateAsync()
    {
        // 台灣銀行佛心提供 CSV 格式，可以直接下載並切割，不需要爬蟲 HTML
        string url = "https://rate.bot.com.tw/xrt/flcsv/0/day";

        var csvData = await _httpClient.GetStringAsync(url);
        var lines = csvData.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);

        foreach (var line in lines)
        {
            var columns = line.Split(',');

            // CSV 第一欄為幣別代碼
            if (columns.Length > 13 && columns[0] == "PHP")
            {
                // 根據台銀 CSV 結構：
                // 索引 2：本行現金買入
                // 索引 12：本行現金賣出
                return (BuyCash: columns[2], SellCash: columns[12]);
            }
        }

        throw new Exception("無法解析台灣銀行披索匯率");
    }
}
